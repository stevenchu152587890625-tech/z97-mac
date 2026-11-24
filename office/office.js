import { readCalendar, readContacts } from './interchange.js'
import { readExcel, readWord } from './ooxml.js'
import { nextOccurrence, occursOn } from './calendar.js'
import { allRecords, exportBackup, makeRecord, maxFileSize, merge, parseBackup, remove, save, statuses, types, watch } from './storage.js'
import { button, recordEditor } from './editors.js'
import { download, element as el, localDate, writeCalendar } from './formats.js'
import { parseDelimited, writeCSV } from './grid.js'
import { legacyRecords } from './legacy.js'

const main = document.getElementById('main')
const dialog = document.getElementById('dialog')
const notice = document.getElementById('notice')
const picker = document.getElementById('import-input')
const desktopLink = document.getElementById('desktop-link')
const initialView = location.hash.slice(1).split('/')[0]
const embedded = window.parent !== window && new URLSearchParams(location.search).has('app')
if (embedded) document.documentElement.classList.add('embedded')
const localStandalone = location.hostname === '127.0.0.1' && location.port === '8767'
if (localStandalone) desktopLink.href = 'http://127.0.0.1:8765/'
else if (location.pathname === '/' || location.pathname === '/index.html') desktopLink.hidden = true
const labels = { home: '總覽', doc: '文件', sheet: '試算表', task: '任務', event: '行事曆',
  contact: '聯絡人', asset: '檔案', trash: '垃圾桶', settings: '備份與設定' }
const icons = { doc: '▤', sheet: '▦', task: '☑', event: '◫', contact: '♙', asset: '▱' }
const localFormatter = new Intl.DateTimeFormat('zh-TW', { dateStyle: 'medium' })
const monthFormatter = new Intl.DateTimeFormat('zh-TW', { year: 'numeric', month: 'long' })
const updatedOrder = (a, b) => b.updated - a.updated
const dueOrder = (a, b) => (a.due || '9999').localeCompare(b.due || '9999') ||
  (a.time || '').localeCompare(b.time || '') || updatedOrder(a, b)

let records = []
let activeEditor
let route = { view: 'home' }
let routeVersion = 0
let projectFilter = ''
let taskMode = 'board'
let taskScope = 'any'
let listFilter = ''
let pendingImport
let noticeTimer
let calendarMonth = new Date()
let focusedDay = localDate()

function notify(text) {
  clearTimeout(noticeTimer)
  notice.textContent = text
  notice.classList.add('visible')
  noticeTimer = setTimeout(() => notice.classList.remove('visible'), 4000)
}

function errorMessage(error) {
  if (error?.name === 'QuotaExceededError') return '瀏覽器儲存空間不足。請先匯出備份，再清理舊附件。'
  return error?.message || String(error)
}

async function guarded(action) {
  try { return await action() } catch (error) { notify(errorMessage(error)) }
}

function parseRoute() {
  const [view, rawId] = location.hash.slice(1).split('/')
  return Object.hasOwn(labels, view) ? { view, id: rawId || '' } : { view: 'home' }
}

function locationFor(view, id) {
  return '#' + view + (id ? '/' + id : '')
}

function visit(view, id) {
  const target = locationFor(view, id)
  if (location.hash === target) syncRoute()
  else location.hash = target
}

function textLink(text, view, id) {
  return el('a', { href: locationFor(view, id) }, text)
}

function liveRecords() {
  return records.filter(record => !record.deleted)
}

function projects() {
  return [...new Set(liveRecords().map(record => record.project).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh-Hant'))
}

const belongs = record => !projectFilter || record.project === projectFilter

function excerpt(record) {
  if (record.kind === 'contact') return [record.company, record.email, record.phone].filter(Boolean).join(' · ')
  if (record.kind === 'sheet') return (record.rows?.length ?? 0) + ' 列'
  if (record.kind === 'asset') return (record.blob?.size ?? 0).toLocaleString() + ' bytes'
  if (record.due) return record.due + (record.time ? ' ' + record.time : '')
  return (record.body ?? '').replace(/^#+\s*/gm, '').replace(/\s+/g, ' ').trim().slice(0, 100)
}

function recordRow(record, actions = []) {
  const label = record.title.trim() || '未命名' + types[record.kind]
  const content = el('div', { class: 'record-cell' },
    el('span', { class: 'record-icon kind-' + record.kind, 'aria-hidden': 'true' }, icons[record.kind]),
    el('span', { class: 'record-copy' }, el('strong', {}, label),
      el('small', {}, [record.project, excerpt(record)].filter(Boolean).join(' · ') || types[record.kind])))
  const control = record.deleted ? el('div', { class: 'record-link' }, content)
    : el('a', { class: 'record-link', href: locationFor(record.kind, record.id) }, content)
  return el('li', { class: 'record-row' }, control,
    el('span', { class: 'record-date' }, localFormatter.format(new Date(record.updated))),
    actions)
}

function recordList(items, empty, limit = Infinity) {
  return items.length ? el('ul', { class: 'record-list' }, items.slice(0, limit).map(record => recordRow(record)))
    : el('p', { class: 'empty' }, empty)
}

function section(title, contents, action) {
  return el('section', { class: 'panel section-panel' },
    el('div', { class: 'section-head' }, el('h2', {}, title), action), contents)
}

function pageHead(title, actions = [], eyebrow, detail) {
  return el('div', { class: 'page-head' },
    el('div', {}, eyebrow ? el('span', { class: 'overline' }, eyebrow) : null, el('h1', {}, title), detail),
    el('div', { class: 'head-actions' }, actions))
}

function selectProject() {
  const chooser = el('select', { 'aria-label': '篩選專案', onchange: event => {
    projectFilter = event.target.value
    render()
  } }, el('option', { value: '' }, '所有專案'), projects().map(project => el('option', { value: project }, project)))
  if (projects().includes(projectFilter)) chooser.value = projectFilter
  else projectFilter = ''
  return chooser
}

function dashboard() {
  const filter = selectProject()
  const items = liveRecords().filter(belongs)
  const today = localDate()
  const pending = items.filter(record => record.kind === 'task' && record.status !== 'done')
  const due = pending.filter(record => record.due && record.due <= today).sort(dueOrder)
  const upcoming = items.filter(record => record.kind === 'event' && record.due)
    .map(record => ({ ...record, due: nextOccurrence(record, today) }))
    .filter(record => record.due).sort(dueOrder)
  const cards = [
    ['待辦任務', pending.length, 'task', 'open'], ['今日及逾期', due.length, 'task', 'due'],
    ['文件及試算表', items.filter(record => ['doc', 'sheet'].includes(record.kind)).length, 'doc'],
    ['已存檔案', items.filter(record => record.kind === 'asset').length, 'asset'],
  ]
  return el('div', { class: 'page' },
    pageHead('工作總覽', [filter, button('建立工作項目', showCreate, 'primary')], null,
      el('p', { class: 'page-detail' }, new Intl.DateTimeFormat('zh-TW', { dateStyle: 'full' }).format(new Date()))),
    el('div', { class: 'metrics' }, cards.map(([label, count, view, scope]) =>
      el('a', { class: 'metric', href: locationFor(view), onclick: () => { if (scope) taskScope = scope } }, el('span', {}, label), el('strong', {}, String(count))))),
    el('div', { class: 'overview-grid' },
      section('最近編輯', recordList(items.sort(updatedOrder), '新增或匯入工作項目後，會顯示在這裡。', 8)),
      el('div', { class: 'stack' },
        section('今日及逾期', recordList(due, '沒有到期任務。', 6), textLink('查看任務', 'task')),
        section('即將到來', recordList(upcoming, '沒有已排定事件。', 6), textLink('開啟行事曆', 'event')))))
}

function kindList(kind) {
  const filter = selectProject()
  const items = liveRecords().filter(record => record.kind === kind && belongs(record)).sort(updatedOrder)
  const results = el('div')
  const search = el('input', { type: 'search', placeholder: '篩選' + types[kind], 'aria-label': '篩選' + types[kind],
    oninput: event => { listFilter = event.target.value; update() } })
  search.value = listFilter
  const update = () => results.replaceChildren(recordList(items.filter(record => searchable(record).includes(listFilter.toLowerCase().trim())),
    listFilter.trim() ? '沒有符合的項目。' : '目前沒有' + types[kind] + '。'))
  update()
  const actions = [filter]
  if (['doc', 'sheet', 'contact', 'asset'].includes(kind)) actions.push(button('匯入', () => askImport(kind === 'asset' ? 'asset' : 'files')))
  if (kind !== 'asset') actions.push(button('＋ 新增' + types[kind], () => create(kind), 'primary'))
  const zone = section(types[kind] + '清單', results, search)
  if (kind === 'asset') {
    zone.addEventListener('dragover', event => {
      if (!event.dataTransfer.types.includes('Files')) return
      event.preventDefault()
      zone.classList.add('drop-active')
    })
    zone.addEventListener('dragleave', event => {
      if (!zone.contains(event.relatedTarget)) zone.classList.remove('drop-active')
    })
    zone.addEventListener('drop', event => {
      event.preventDefault()
      zone.classList.remove('drop-active')
      importFiles([...event.dataTransfer.files], 'asset')
    })
  }
  return el('div', { class: 'page' }, pageHead(labels[kind], actions), zone)
}

async function setStatus(id, status) {
  await guarded(async () => {
    const record = records.find(entry => entry.id === id && !entry.deleted && entry.kind === 'task')
    if (!record || record.status === status || !Object.hasOwn(statuses, status)) return
    await save({ ...record, status })
    await refresh()
    render()
  })
}

function taskCard(record) {
  const status = record.status || 'todo'
  const marker = record.due ? (record.due < localDate() && status !== 'done' ? '逾期 ' : '') + record.due : ''
  return el('article', {
    class: 'task-card', draggable: true,
    ondragstart: event => {
      event.dataTransfer.effectAllowed = 'move'
      event.dataTransfer.setData('application/x-z97-task', record.id)
    },
  },
  el('a', { class: 'task-title', href: locationFor('task', record.id) }, record.title || '未命名任務'),
  el('div', { class: 'task-meta' },
    record.priority === 'high' ? el('span', { class: 'tag urgent' }, '高優先') : null,
    marker ? el('span', {}, marker) : null,
    record.project ? el('span', { class: 'tag' }, record.project) : null),
  button(status === 'done' ? '重新開啟' : '標為完成', () => setStatus(record.id, status === 'done' ? 'todo' : 'done'), 'quiet'))
}

function tasks() {
  const filter = selectProject()
  const today = localDate()
  const scopes = { any: '全部日期與狀態', open: '未完成', due: '今日及逾期', upcoming: '未來到期' }
  const withinScope = record => {
    if (taskScope === 'any') return true
    if (record.status === 'done') return false
    if (taskScope === 'due') return record.due && record.due <= today
    if (taskScope === 'upcoming') return record.due && record.due > today
    return true
  }
  const items = liveRecords().filter(record => record.kind === 'task' && belongs(record) && withinScope(record)).sort(dueOrder)
  const chooser = el('select', { 'aria-label': '篩選任務', onchange: event => { taskScope = event.target.value; render() } },
    Object.entries(scopes).map(([key, label]) => el('option', { value: key }, label)))
  chooser.value = taskScope
  const exportTasks = () => {
    if (!items.length) return notify('目前沒有可匯出的任務。')
    const names = { low: '低', normal: '一般', high: '高' }
    const rows = [['名稱', '專案', '狀態', '優先順序', '截止日', '時間', '備註']]
    for (const item of items) rows.push([item.title, item.project, statuses[item.status || 'todo'],
      names[item.priority || 'normal'], item.due, item.time, item.body].map(value => "'" + (value || '')))
    download('任務-' + today + '.csv', writeCSV(rows), 'text/csv;charset=utf-8')
  }
  const viewSwitch = el('div', { class: 'segmented' }, ['board', 'list'].map((mode, index) => {
    const control = button(index ? '清單' : '看板', () => { taskMode = mode; render() }, 'quiet')
    control.setAttribute('aria-pressed', String(taskMode === mode))
    return control
  }))
  const contents = taskMode === 'list' ? section('所有任務', recordList(items, '目前沒有任務。'))
    : el('div', { class: 'board' }, Object.entries(statuses).map(([status, name]) => {
      const inLane = items.filter(record => (record.status || 'todo') === status)
      return el('section', {
        class: 'lane',
        ondragover: event => {
          if (!event.dataTransfer.types.includes('application/x-z97-task')) return
          event.preventDefault()
          event.dataTransfer.dropEffect = 'move'
          event.currentTarget.classList.add('drop-active')
        },
        ondragleave: event => {
          if (!event.currentTarget.contains(event.relatedTarget)) event.currentTarget.classList.remove('drop-active')
        },
        ondrop: event => {
          event.preventDefault()
          event.currentTarget.classList.remove('drop-active')
          setStatus(event.dataTransfer.getData('application/x-z97-task'), status)
        },
      }, el('div', { class: 'lane-head' }, el('h2', {}, name), el('span', {}, String(inLane.length))),
      inLane.length ? inLane.map(taskCard) : el('p', { class: 'lane-empty' }, '沒有任務'))
    }))
  return el('div', { class: 'page' },
    pageHead('任務', [filter, chooser, viewSwitch, button('下載 CSV', exportTasks), button('＋ 新增任務', () => create('task'), 'primary')]), contents)
}

function calendar() {
  const filter = selectProject()
  const items = liveRecords().filter(record => (record.kind === 'event' || record.kind === 'task') && record.due && belongs(record))
  const today = localDate()
  const year = calendarMonth.getFullYear()
  const month = calendarMonth.getMonth()
  const first = new Date(year, month, 1)
  const offset = (first.getDay() + 6) % 7
  const days = el('div', { class: 'calendar-grid' },
    ['一', '二', '三', '四', '五', '六', '日'].map(name => el('div', { class: 'weekday' }, name)))
  for (let index = 0; index < 42; index++) {
    const date = new Date(year, month, index + 1 - offset)
    const day = localDate(date)
    const dayItems = items.filter(record => occursOn(record, day))
    const cell = button('', () => { focusedDay = day; calendarMonth = date; render() }, 'day')
    cell.classList.toggle('other-month', date.getMonth() !== month)
    cell.classList.toggle('today', day === today)
    cell.classList.toggle('selected-day', day === focusedDay)
    cell.setAttribute('aria-label', day + '，' + dayItems.length + ' 筆')
    cell.setAttribute('aria-pressed', String(day === focusedDay))
    cell.append(el('strong', {}, String(date.getDate())),
      ...dayItems.slice(0, 2).map(item => el('span', { class: 'day-item kind-' + item.kind }, item.title || types[item.kind])),
      dayItems.length > 2 ? el('small', {}, '+' + (dayItems.length - 2)) : '')
    days.append(cell)
  }
  const monthChange = offset => {
    calendarMonth = new Date(year, month + offset, 1)
    focusedDay = localDate(calendarMonth)
    render()
  }
  return el('div', { class: 'page' },
    pageHead('行事曆', [filter, button('匯入 ICS', () => askImport('files')), button('匯出 ICS', () => {
      const events = items.filter(record => record.kind === 'event')
      if (!events.length) return notify('目前沒有可匯出的行程。')
      download('行事曆.ics', writeCalendar(events), 'text/calendar;charset=utf-8')
    }), button('＋ 新增事件', () => create('event', { due: focusedDay }), 'primary')]),
    el('div', { class: 'calendar-layout' },
      el('section', { class: 'panel calendar-panel' },
        el('div', { class: 'section-head' }, el('h2', {}, monthFormatter.format(first)),
          el('div', { class: 'toolbar' }, button('上一月', () => monthChange(-1), 'quiet'),
            button('今天', () => { calendarMonth = new Date(); focusedDay = today; render() }, 'quiet'),
            button('下一月', () => monthChange(1), 'quiet'))), days),
      section(focusedDay, recordList(items.filter(record => occursOn(record, focusedDay)).sort(dueOrder), '這一天沒有安排。'))))
}

function trash() {
  const items = records.filter(record => record.deleted).sort(updatedOrder)
  const list = items.length ? el('ul', { class: 'record-list' }, items.map(record =>
    recordRow(record, [
      button('還原', () => guarded(async () => {
        await save({ ...record, deleted: false })
        await refresh()
        render()
        notify('已還原。')
      })),
      button('永久刪除', () => confirmAction('永久刪除', '確定要永久刪除「' + record.title + '」？', async () => {
        await remove(record)
        await refresh()
        render()
      }), 'danger'),
    ]))) : el('p', { class: 'empty' }, '垃圾桶是空的。')
  return el('div', { class: 'page' }, pageHead('垃圾桶'), section('已移除項目', list))
}

function settings() {
  const storageText = el('p', { class: 'muted' }, '工作台資料保存在此瀏覽器的 IndexedDB。換瀏覽器、網址或裝置時，請用備份匯入。')
  const estimate = el('p', { class: 'muted' })
  navigator.storage?.estimate?.().then(info => {
    if (info.usage !== undefined) estimate.textContent = '此網站目前使用約 ' + (info.usage / 1024 / 1024).toFixed(1) + ' MB 儲存空間。'
  }).catch(() => {})
  return el('div', { class: 'page' }, pageHead('備份與設定'),
    section('資料備份', el('div', { class: 'form-panel' }, storageText, estimate,
      el('div', { class: 'toolbar' },
        button('下載完整備份', () => guarded(async () => {
          download('工作台備份-' + localDate() + '.json', await exportBackup(), 'application/json')
          notify('已產生備份。')
        }), 'primary'),
        button('還原備份', () => askImport('backup')),
        button('請求保留瀏覽器儲存', () => guarded(async () => {
          const granted = await navigator.storage?.persist?.()
          notify(granted ? '瀏覽器已允許持久儲存。' : '瀏覽器未授權。請定期下載備份。')
        }))),
      el('p', { class: 'muted' }, '備份包含附件、公式與垃圾桶，可合併進來。同一 ID 的現有記錄優先保留。'))),
    section('帶入舊桌面資料', el('div', { class: 'form-panel' },
      el('p', { class: 'muted' }, '可在相同網址讀取原桌面使用此瀏覽器保存的筆記、提醒、行程、通訊錄和文字檔。先選取要複製的項目。'),
      button('檢視可匯入的項目', showLegacy),
      localStandalone ? el('p', { class: 'muted' }, '如要帶入原桌面的內容，先在 ',
        el('a', { href: 'http://127.0.0.1:8765/office/#settings' }, '原桌面旁的工作台'),
        ' 複製選取項目、下載備份，再回這裡還原。') : null,
      el('p', { class: 'muted' }, '可匯入 TXT、Markdown、DOCX、CSV、TSV、XLSX、ICS 和 vCard。Word 保留文字與基本段落，Excel 匯入第一張表的值與支援的公式。可在檔案區另存原件。'),
      button('從裝置匯入檔案', () => askImport('files')))))
}

function showDialog(title, content) {
  dialog.replaceChildren(el('div', { class: 'dialog-bar' },
    el('h2', { id: 'dialog-title' }, title), button('關閉', () => dialog.close(), 'quiet')), content)
  if (!dialog.open) dialog.showModal()
}

function confirmAction(title, message, action) {
  showDialog(title, el('div', { class: 'dialog-body' },
    el('p', {}, message), el('div', { class: 'toolbar' },
      button('取消', () => dialog.close()),
      button(title, () => guarded(async () => { await action(); dialog.close() }), 'danger'))))
}

function searchable(record) {
  return [record.title, record.project, record.body, record.company, record.email, record.phone,
    record.filename, record.due, types[record.kind], record.rows?.map(row => row.join(' ')).join(' ')]
    .filter(Boolean).join(' ').toLowerCase()
}

function showSearch() {
  const items = liveRecords().filter(record => !embedded || !Object.hasOwn(types, initialView) || record.kind === initialView).sort(updatedOrder)
  const cache = new Map(items.map(record => [record.id, searchable(record)]))
  const result = el('div', { class: 'search-results' })
  const input = el('input', { autofocus: true, type: 'search', placeholder: '輸入關鍵字或專案名稱…',
    'aria-label': '搜尋工作資料', oninput: update, onkeydown: event => {
      if (event.key === 'Enter') result.querySelector('a')?.click()
    } })
  function update() {
    const words = input.value.trim().toLowerCase().split(/\s+/).filter(Boolean)
    result.replaceChildren(recordList(items.filter(record => words.every(word => cache.get(record.id).includes(word))),
      '沒有符合的工作資料。', 80))
  }
  result.addEventListener('click', event => { if (event.target.closest('a')) dialog.close() })
  showDialog('搜尋', el('div', { class: 'dialog-body' }, input, result))
  update()
  input.focus()
}

const templates = [
  ['doc', '會議紀錄', { title: '會議紀錄', body: '## 會議資訊\n- 日期：\n- 參與者：\n\n## 議程\n- \n\n## 決議\n- \n\n## 後續行動\n- 事項／負責人／截止日：\n' }],
  ['doc', '工作週報', { title: '工作週報', body: '## 本週成果\n- \n\n## 進行中\n- \n\n## 風險與待協助事項\n- \n\n## 下週計畫\n- \n' }],
  ['sheet', '費用紀錄', { title: '費用紀錄', rows: [
    ['日期', '類別', '品項', '金額', '備註'],
    ...Array.from({ length: 10 }, () => ['', '', '', '', '']),
    ['合計', '', '', '=SUM(D2:D11)', ''],
  ] }],
  ['sheet', '工時紀錄', { title: '工時紀錄', rows: [
    ['日期', '人員', '工作內容', '工時', '費率', '小計'],
    ...Array.from({ length: 10 }, (_, index) => ['', '', '', '', '', '=D' + (index + 2) + '*E' + (index + 2)]),
    ['合計', '', '', '=SUM(D2:D11)', '', '=SUM(F2:F11)'],
  ] }],
]

function showCreate() {
  showDialog('新增工作項目', el('div', { class: 'dialog-body create-options' },
    Object.entries(types).map(([kind, label]) => button(icons[kind] + '　' + label, () => {
      dialog.close()
      if (kind === 'asset') askImport('asset')
      else create(kind)
    })), button('匯入文件或表格', () => { dialog.close(); askImport('files') }),
    el('h3', { class: 'create-heading' }, '從範本開始'),
    templates.map(([kind, label, fields]) => button(label, () => {
      dialog.close()
      create(kind, structuredClone(fields))
    }))))
}

function showLegacy() {
  const existing = new Set(records.map(record => record.id))
  const entries = legacyRecords().filter(entry => !existing.has(entry.record.id))
  const selections = new Set()
  const selectedCount = el('span', { class: 'muted' }, '尚未選取')
  const choices = el('div', { class: 'legacy-choices' }, entries.map(({ record, source }) => {
    const input = el('input', { type: 'checkbox', onchange: event => {
      if (event.target.checked) selections.add(record.id)
      else selections.delete(record.id)
      selectedCount.textContent = '已選取 ' + selections.size + ' 筆'
    } })
    return el('label', {}, input, el('span', {}, record.title), el('small', {}, source))
  }))
  const select = status => {
    for (const input of choices.querySelectorAll('input')) {
      if (input.checked !== status) input.click()
    }
  }
  showDialog('從舊桌面帶入', el('div', { class: 'dialog-body' },
    el('p', { class: 'muted' }, '原桌面的示範資料也可能出現。只勾選你要保留的項目。'),
    entries.length ? choices : el('p', { class: 'empty' }, '目前找不到可帶入的新項目。請確認舊資料和工作台在同一個網址。'),
    el('div', { class: 'toolbar' },
      button('全選', () => select(true), 'quiet', !entries.length),
      button('清除選取', () => select(false), 'quiet', !entries.length), selectedCount,
      button('複製選取項目', () => guarded(async () => {
        if (!selections.size) return notify('請先選擇要複製的項目。')
        const count = await merge(entries.filter(entry => selections.has(entry.record.id)).map(entry => entry.record))
        await refresh()
        dialog.close()
        notify('已帶入 ' + count + ' 筆。')
      }), 'primary', !entries.length))))
}

async function create(kind, fields = {}) {
  await guarded(async () => {
    if (activeEditor && !await flush(activeEditor)) return
    const initial = { doc: { body: '' }, sheet: { rows: Array.from({ length: 12 }, () => Array(6).fill('')) },
      task: { status: 'todo', priority: 'normal', due: '', body: '' }, event: { due: localDate(), body: '' },
      contact: { email: '', phone: '', company: '', body: '' } }
    const stored = await save(makeRecord(kind, { project: projectFilter, ...initial[kind], ...fields }))
    await refresh()
    visit(kind, stored.id)
  })
}

function askImport(mode) {
  pendingImport = mode
  picker.multiple = mode !== 'backup'
  picker.accept = mode === 'backup' ? '.json,application/json' : ''
  picker.value = ''
  picker.click()
}

async function importFiles(files, mode) {
  if (!files.length) return
  await guarded(async () => {
    if (activeEditor && !await flush(activeEditor)) return
    if (mode === 'backup') {
      const entries = parseBackup(await files[0].text())
      confirmAction('還原備份', '備份包含 ' + entries.length + ' 筆資料。同一 ID 的現有項目會保留。', async () => {
        const count = await merge(entries)
        await refresh()
        render()
        notify('已還原 ' + count + ' 筆。')
      })
      return
    }
    const entries = []
    const issues = []
    const details = []
    for (const file of files) {
      const title = file.name.slice(0, 250)
      const ext = file.name.split('.').pop().toLowerCase()
      try {
        if (file.size > maxFileSize) throw new Error('檔案超過 25 MB。')
        if (mode !== 'asset' && ['ics', 'vcf', 'vcard'].includes(ext)) {
          const isCalendar = ext === 'ics'
          const parsed = (isCalendar ? readCalendar : readContacts)(await file.text())
          for (const item of parsed.entries) entries.push(makeRecord(isCalendar ? 'event' : 'contact', { project: projectFilter, ...item }))
          if (parsed.skipped) details.push(title + '：略過 ' + parsed.skipped + ' 個不支援或已取消的項目。')
          if (parsed.simplified) details.push(title + '：' + parsed.simplified + ' 筆的複雜重複規則或長度已簡化。')
        } else if (mode !== 'asset' && ext === 'docx') {
          const body = await readWord(file)
          entries.push(makeRecord('doc', { title, project: projectFilter, body }))
          details.push(title + '：已帶入文字與基本格式。')
        } else if (mode !== 'asset' && ext === 'xlsx') {
          const { rows, sheetCount } = await readExcel(file)
          entries.push(makeRecord('sheet', { title, project: projectFilter, rows }))
          if (sheetCount > 1) details.push(title + '：已匯入第一張工作表（原檔共 ' + sheetCount + ' 張）。')
        } else if (mode !== 'asset' && ['csv', 'tsv'].includes(ext)) {
          entries.push(makeRecord('sheet', { title, project: projectFilter, rows: parseDelimited(await file.text(), ext === 'tsv' ? '\t' : ',') }))
        } else if (mode !== 'asset' && ['txt', 'md', 'markdown', 'log'].includes(ext)) {
          const body = await file.text()
          if (body.length > 2_000_000) throw new Error('文件超過 200 萬字元，可改從檔案頁匯入為附件。')
          entries.push(makeRecord('doc', { title, project: projectFilter, body }))
        } else {
          entries.push(makeRecord('asset', { title, project: projectFilter, filename: title,
            mime: file.type, blob: file }))
        }
      } catch (error) {
        issues.push(title + '：' + errorMessage(error))
      }
    }
    const count = await merge(entries)
    await refresh()
    if (!issues.length) {
      if (entries.length === 1) visit(entries[0].kind, entries[0].id)
      else if (count && entries.every(record => record.kind === entries[0].kind)) visit(entries[0].kind)
      else if (count) visit('home')
      notify(['已匯入 ' + count + ' 筆。', ...details].join(' '))
    } else {
      if (!activeEditor) render()
      showDialog('匯入結果', el('div', { class: 'dialog-body' },
        el('p', {}, '已匯入 ' + count + ' 筆。以下檔案需要處理。'),
        el('ul', { class: 'import-errors' }, issues.map(issue => el('li', {}, issue)))))
    }
  })
}

async function refresh() {
  records = await allRecords()
}

function flush(editor) {
  if (!editor) return Promise.resolve(true)
  clearTimeout(editor.timer)
  if (editor.pending) return editor.pending
  if (editor.error) return Promise.resolve(false)
  if (!editor.dirty) return Promise.resolve(true)
  editor.pending = (async () => {
    try {
      while (editor.dirty) {
        const { draft, generation } = editor
        editor.indicator.textContent = '儲存中…'
        const stored = await save(draft, editor.saved.rev)
        editor.saved = stored
        editor.draft = { ...editor.draft, rev: stored.rev, updated: stored.updated }
        editor.dirty = generation !== editor.generation
        const index = records.findIndex(record => record.id === stored.id)
        if (index >= 0) records[index] = stored
        editor.updateRelated(stored)
      }
      editor.indicator.textContent = '已儲存'
      return true
    } catch (error) {
      editor.error = error
      editor.indicator.replaceChildren(el('span', {}, errorMessage(error)),
        button('保留草稿為副本', () => guarded(async () => {
          if (activeEditor !== editor) return
          const { id, rev, ...fields } = editor.draft
          const stored = await save(makeRecord(fields.kind, { ...fields, title: fields.title + '（副本）' }))
          editor.error = undefined
          editor.dirty = false
          await refresh()
          visit(stored.kind, stored.id)
        }), 'quiet'),
        button('捨棄草稿並重新載入', () => confirmAction('捨棄草稿', '尚未儲存的變更會被捨棄。', async () => {
          editor.error = undefined
          editor.dirty = false
          await refresh()
          render()
        }), 'quiet'),
        error.code !== 'CONFLICT' ? button('重試', () => {
          editor.error = undefined
          flush(editor)
        }, 'quiet') : null)
      return false
    }
  })().finally(() => { editor.pending = undefined })
  return editor.pending
}

function openRecord(record) {
  const urls = new Set()
  const related = el('div')
  const updateRelated = item => {
    const others = item.project ? liveRecords().filter(candidate => candidate.id !== item.id && candidate.project === item.project && (!embedded || candidate.kind === item.kind))
      .sort(updatedOrder) : []
    related.hidden = !others.length
    related.replaceChildren(...(others.length ? [section('同一專案的其他項目', recordList(others, '', 8))] : []))
  }
  updateRelated(record)
  const indicator = el('div', { class: 'save-state', role: 'status', 'aria-live': 'polite' }, '已儲存')
  const editor = {
    saved: record, draft: record, generation: 0, dirty: false, indicator,
    projects: projects(), notify, updateRelated,
    objectURL(blob) {
      const url = URL.createObjectURL(blob)
      urls.add(url)
      return url
    },
    dispose() {
      clearTimeout(editor.timer)
      for (const url of urls) URL.revokeObjectURL(url)
    },
    change(patch) {
      editor.draft = { ...editor.draft, ...patch }
      editor.generation++
      editor.dirty = true
      if (editor.error) return
      indicator.textContent = '尚未儲存…'
      clearTimeout(editor.timer)
      editor.timer = setTimeout(() => flush(editor), 450)
    },
  }
  activeEditor = editor
  return el('div', { class: 'page editor-page' },
    el('div', { class: 'editor-actions' },
      textLink('← ' + labels[record.kind], record.kind),
      el('div', { class: 'toolbar' },
        button('立即儲存', () => flush(editor)),
        button('建立副本', () => guarded(async () => {
          if (!await flush(editor)) return
          const { id, rev, updated, created, ...fields } = editor.saved
          const title = fields.title.slice(0, 245) + '（副本）'
          const copy = await save(makeRecord(fields.kind, { ...fields, title }))
          await refresh()
          visit(copy.kind, copy.id)
          notify('已建立副本。')
        })),
        button('移至垃圾桶', () => guarded(async () => {
          if (!await flush(editor)) return
          await save({ ...editor.saved, deleted: true })
          await refresh()
          visit(record.kind)
          notify('已移至垃圾桶。')
        }), 'quiet'))),
    recordEditor(record, editor), related)
}

function render() {
  for (const link of document.querySelectorAll('[data-view]')) {
    if (link.dataset.view === route.view) link.setAttribute('aria-current', 'page')
    else link.removeAttribute('aria-current')
  }
  if (activeEditor) activeEditor.dispose()
  activeEditor = undefined
  const found = route.id && records.find(record => record.id === route.id && record.kind === route.view && !record.deleted)
  let view
  if (found) view = openRecord(found)
  else if (route.view === 'home') view = dashboard()
  else if (route.view === 'task') view = tasks()
  else if (route.view === 'event') view = calendar()
  else if (route.view === 'trash') view = trash()
  else if (route.view === 'settings') view = settings()
  else view = kindList(route.view)
  main.replaceChildren(view)
  document.title = (found?.title || labels[route.view]) + ' · 工作台'
}

async function syncRoute() {
  const change = ++routeVersion
  const requested = parseRoute()
  const editor = activeEditor
  // A blur in another control can schedule spreadsheet recalculation before this switch.
  if (editor && !await flush(editor)) {
    if (change === routeVersion) history.replaceState(null, '', locationFor(route.view, route.id))
    return
  }
  if (change !== routeVersion) return
  if (route.view !== requested.view) listFilter = ''
  route = requested
  render()
  if (!embedded || document.hasFocus()) main.focus()
}

picker.addEventListener('change', () => importFiles([...picker.files], pendingImport))
document.getElementById('search-button').addEventListener('click', showSearch)
document.getElementById('new-button').addEventListener('click', showCreate)
window.addEventListener('hashchange', syncRoute)

function runDesktopCommand(command) {
  const invokeButton = label => {
    const control = Array.from(main.querySelectorAll('button')).find(button => button.textContent.trim() === label)
    if (control) control.click()
    else notify('請先開啟要操作的項目。')
  }
  const actions = {
    new: () => Object.hasOwn(types, initialView) && initialView !== 'asset'
      ? create(initialView, initialView === 'event' ? { due: focusedDay } : {}) : showCreate(),
    search: showSearch,
    save: () => activeEditor ? flush(activeEditor) : notify('目前沒有開啟的項目。'),
    import: () => askImport(initialView === 'asset' ? 'asset' : 'files'),
    word: () => invokeButton('下載 Word'),
    markdown: () => invokeButton('下載 Markdown'),
    print: () => invokeButton('列印／PDF'),
    ics: () => invokeButton(activeEditor ? '下載行事曆 ICS' : '匯出 ICS'),
    vcard: () => invokeButton('下載 vCard'),
    backup: () => invokeButton('下載完整備份'),
    restore: () => askImport('backup'),
    today: async () => {
      if (activeEditor && !await flush(activeEditor)) return
      calendarMonth = new Date()
      focusedDay = localDate()
      visit('event')
    },
  }
  return Object.hasOwn(actions, command) ? actions[command]() : undefined
}

if (embedded) window.z97Desktop = {
  command: runDesktopCommand,
  hasPending: () => !!activeEditor?.dirty,
  flush: () => flush(activeEditor),
}
window.addEventListener('keydown', event => {
  if (!(event.ctrlKey || event.metaKey) || event.altKey) return
  if (event.key.toLowerCase() === 'k') { event.preventDefault(); showSearch() }
  if (event.key.toLowerCase() === 's' && activeEditor) { event.preventDefault(); flush(activeEditor) }
  if (embedded) {
    const commands = { n: 'new', f: 'search', p: 'print', o: 'import' }
    const key = event.key.toLowerCase()
    if (Object.hasOwn(commands, key) && (key !== 'p' || initialView === 'doc')) {
      event.preventDefault()
      runDesktopCommand(commands[key])
    }
  }
})
window.addEventListener('beforeunload', event => {
  if (activeEditor?.dirty) { event.preventDefault(); event.returnValue = '' }
})
watch(async () => {
  await guarded(async () => {
    await refresh()
    if (activeEditor) activeEditor.updateRelated(activeEditor.saved)
    else if (!dialog.open) render()
  })
})

async function init() {
  try {
    await refresh()
    await syncRoute()
  } catch (error) {
    main.replaceChildren(el('div', { class: 'page' }, pageHead('無法開啟工作台'),
      el('p', {}, errorMessage(error)), button('重新載入', () => location.reload())))
  }
}
init()

if ('serviceWorker' in navigator && window.isSecureContext) {
  navigator.serviceWorker.register('./sw.js').catch(() => {})
}
