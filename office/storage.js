const DB_NAME = 'z97-office'
const STORE = 'records'
const channel = typeof BroadcastChannel === 'function' ? new BroadcastChannel(DB_NAME) : null
let connection

export const types = {
  doc: '文件', sheet: '試算表', task: '任務', event: '事件',
  contact: '聯絡人', asset: '檔案',
}
export const statuses = { todo: '待處理', doing: '進行中', done: '完成' }
export const maxFileSize = 25 * 1024 * 1024
export const maxRows = 1000
export const maxColumns = 40

export function makeRecord(kind, fields = {}) {
  const id = typeof crypto.randomUUID === 'function'
    ? crypto.randomUUID()
    : Array.from(crypto.getRandomValues(new Uint8Array(16)), n => n.toString(16).padStart(2, '0')).join('')
  return {
    id, kind, title: '未命名' + types[kind], project: '', created: Date.now(),
    updated: Date.now(), rev: 0, deleted: false, ...fields,
  }
}

export function openDB() {
  if (connection) return connection
  connection = new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1)
    request.onupgradeneeded = () => request.result.createObjectStore(STORE, { keyPath: 'id' })
    request.onsuccess = () => {
      const db = request.result
      db.onversionchange = () => {
        db.close()
        connection = undefined
      }
      resolve(db)
    }
    request.onerror = () => reject(request.error)
    request.onblocked = () => reject(new Error('請關閉其他工作台頁籤，再重新載入。'))
  }).catch(error => {
    connection = undefined
    throw error
  })
  return connection
}

async function transaction(mode, run) {
  const db = await openDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, mode)
    let result
    let failure
    const abort = error => {
      failure = error
      tx.abort()
    }
    tx.oncomplete = () => resolve(result)
    tx.onabort = () => reject(failure || tx.error || new Error('資料寫入已取消。'))
    try {
      run(tx.objectStore(STORE), value => { result = value }, abort)
    } catch (error) {
      abort(error)
    }
  })
}

export function allRecords() {
  return transaction('readonly', (store, finish) => {
    const request = store.getAll()
    request.onsuccess = () => finish(request.result)
  })
}

function conflict(existing) {
  return Object.assign(new Error('另一個頁籤已修改這筆資料。可重新載入，或把目前草稿另存副本。'), { existing, code: 'CONFLICT' })
}

export async function save(record, expected = record.rev) {
  const stored = await transaction('readwrite', (store, finish, abort) => {
    const request = store.get(record.id)
    request.onsuccess = () => {
      const current = request.result
      if ((current?.rev ?? 0) !== expected) return abort(conflict(current))
      const next = { ...record, rev: expected + 1, updated: Date.now() }
      store.put(next)
      finish(next)
    }
  })
  channel?.postMessage(stored.id)
  return stored
}

export async function remove(record) {
  await transaction('readwrite', (store, finish, abort) => {
    const request = store.get(record.id)
    request.onsuccess = () => {
      if (request.result?.rev !== record.rev) return abort(conflict(request.result))
      store.delete(record.id)
      finish(true)
    }
  })
  channel?.postMessage(record.id)
}

export async function merge(records) {
  const count = await transaction('readwrite', (store, finish) => {
    const current = store.getAllKeys()
    current.onsuccess = () => {
      const ids = new Set(current.result)
      let count = 0
      for (const entry of records) {
        if (ids.has(entry.id)) continue
        store.add({ ...entry, rev: 1 })
        ids.add(entry.id)
        count++
      }
      finish(count)
    }
  })
  if (count) channel?.postMessage(null)
  return count
}

export function watch(handler) {
  if (channel) channel.onmessage = handler
}

const allowedString = (input, max = 10000) => typeof input === 'string' ? input.slice(0, max) : ''
const backupDate = value => {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return ''
  const date = new Date(value + 'T00:00:00Z')
  return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value ? value : ''
}
const backupTime = value => typeof value === 'string' && /^(?:[01]\d|2[0-3]):[0-5]\d$/.test(value) ? value : ''
const backupStamp = value => Number.isFinite(value) && Math.abs(value) <= 8.64e15 ? value : Date.now()

function asDataURL(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result)
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(blob)
  })
}

export async function exportBackup() {
  const records = await allRecords()
  const serialized = []
  for (const record of records) {
    const { blob, ...fields } = record
    serialized.push(blob ? { ...fields, base64: (await asDataURL(blob)).split(',')[1] } : fields)
  }
  return JSON.stringify({ format: 'z97-office', version: 1, records: serialized })
}

export function parseBackup(text) {
  const data = JSON.parse(text)
  if (data?.format !== 'z97-office' || data.version !== 1 || !Array.isArray(data.records)) {
    throw new Error('檔案格式錯誤。請選擇工作台匯出的 JSON 備份。')
  }
  return data.records.map(entry => {
    if (!entry || !Object.hasOwn(types, entry.kind) || typeof entry.id !== 'string' || !/^[\w:-]{1,250}$/.test(entry.id)) {
      throw new Error('備份包含無效的記錄，未還原任何資料。')
    }
    const record = {
      id: entry.id, kind: entry.kind, title: allowedString(entry.title, 250) || '未命名',
      project: allowedString(entry.project, 120), deleted: entry.deleted === true,
      created: backupStamp(entry.created), updated: backupStamp(entry.updated),
    }
    if (entry.kind === 'doc') record.body = allowedString(entry.body, 2_000_000)
    if (entry.kind === 'task' || entry.kind === 'event') {
      record.due = backupDate(entry.due)
      record.time = backupTime(entry.time)
      record.body = allowedString(entry.body, 100000)
    }
    if (entry.kind === 'event') {
      record.endDate = backupDate(entry.endDate)
      record.endTime = backupTime(entry.endTime)
      record.repeat = ['daily', 'weekly', 'monthly'].includes(entry.repeat) ? entry.repeat : ''
      record.repeatUntil = backupDate(entry.repeatUntil)
    }
    if (entry.kind === 'task') {
      record.status = Object.hasOwn(statuses, entry.status) ? entry.status : 'todo'
      record.priority = ['low', 'normal', 'high'].includes(entry.priority) ? entry.priority : 'normal'
    }
    if (entry.kind === 'contact') {
      for (const key of ['email', 'phone', 'company', 'body']) record[key] = allowedString(entry[key])
    }
    if (entry.kind === 'sheet') {
      if (!Array.isArray(entry.rows) || entry.rows.length > maxRows ||
          entry.rows.some(row => !Array.isArray(row) || row.length > maxColumns)) {
        throw new Error('備份包含無效或過大的試算表。')
      }
      record.rows = entry.rows.map(row => row.map(cell => allowedString(cell, 20000)))
    }
    if (entry.kind === 'asset') {
      if (typeof entry.base64 !== 'string' || entry.base64.length > Math.ceil(maxFileSize / 3) * 4) {
        throw new Error('備份包含無效或大於 25 MB 的附件。')
      }
      const bytes = Uint8Array.from(atob(entry.base64), char => char.charCodeAt(0))
      if (bytes.length > maxFileSize) throw new Error('附件大於 25 MB。')
      record.mime = allowedString(entry.mime, 200)
      record.filename = allowedString(entry.filename, 250) || record.title
      record.blob = new Blob([bytes], { type: record.mime })
    }
    return record
  })
}
