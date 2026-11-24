import { addDays } from './calendar.js'
import { makeRecord } from './storage.js'
import { localDate } from './formats.js'

function load(key) {
  try { return JSON.parse(localStorage.getItem(key) ?? 'null')?.state ?? {} } catch { return {} }
}

function fromTime(timestamp) {
  const date = new Date(timestamp)
  if (typeof timestamp !== 'number' || !Number.isFinite(date.getTime())) return { due: '', time: '' }
  return { due: localDate(date), time: [date.getHours(), date.getMinutes()].map(value => String(value).padStart(2, '0')).join(':') }
}

function noteText(html) {
  const template = document.createElement('template')
  template.innerHTML = html
  const convert = node => {
    if (node.nodeType === Node.TEXT_NODE) return node.textContent
    if (['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEMPLATE'].includes(node.tagName)) return ''
    const text = Array.from(node.childNodes, convert).join('')
    if (/^H[1-6]$/.test(node.tagName)) return '\n' + '#'.repeat(Number(node.tagName[1])) + ' ' + text + '\n'
    if (node.tagName === 'LI') return '- ' + text.trim() + '\n'
    if (node.tagName === 'BR') return '\n'
    if (['P', 'DIV', 'UL', 'OL', 'PRE', 'BLOCKQUOTE'].includes(node.tagName)) return text + '\n'
    if (['STRONG', 'B'].includes(node.tagName) && text.trim()) return '**' + text + '**'
    if (['EM', 'I'].includes(node.tagName) && text.trim()) return '*' + text + '*'
    return text
  }
  return convert(template.content).trim()
}

export function legacyRecords() {
  const entries = []
  const text = (value, max = 100000) => typeof value === 'string' ? value.slice(0, max) : ''
  const add = (source, entry, kind, fields) => {
    if (typeof entry?.id !== 'string' || !/^[\w:-]{1,170}$/.test(entry.id)) return
    const keys = { '備忘錄': 'notes', '提醒事項': 'reminders', '行事曆': 'calendar', '通訊錄': 'contacts', Finder: 'finder' }
    const id = 'legacy:' + keys[source] + ':' + entry.id
    const safeFields = Object.fromEntries(Object.entries(fields).filter(([, value]) => value !== undefined))
    const record = makeRecord(kind, { ...safeFields, id })
    record.title = text(record.title, 250).trim() || '未命名'
    record.project = text(record.project, 120)
    entries.push({ record, source })
  }
  const noteState = load('macos27:app:notes')
  for (const note of Array.isArray(noteState.notes) ? noteState.notes : []) {
    if (!note || note.deleted || typeof note.html !== 'string') continue
    const body = noteText(note.html).slice(0, 2_000_000)
    if (body) add('備忘錄', note, 'doc', { title: body.split('\n')[0].replace(/^#+\s*/, ''), body })
  }
  const reminderState = load('macos27:app:reminders')
  const names = new Map((Array.isArray(reminderState.lists) ? reminderState.lists : []).filter(Boolean).map(list => [list.id, list.name]))
  for (const item of Array.isArray(reminderState.reminders) ? reminderState.reminders : []) {
    if (!item?.title) continue
    add('提醒事項', item, 'task', { title: item.title, project: names.get(item.listId),
      body: text(item.notes), status: item.completed ? 'done' : 'todo',
      priority: item.flagged || item.priority >= 2 ? 'high' : 'normal', ...fromTime(item.due) })
  }
  const calendarState = load('macos27:app:calendar')
  const calendarNames = new Map((Array.isArray(calendarState.calendars) ? calendarState.calendars : []).filter(Boolean).map(item => [item.id, item.name]))
  for (const item of Array.isArray(calendarState.events) ? calendarState.events : []) {
    if (!item?.title) continue
    const start = fromTime(item.start)
    const end = fromTime(item.end)
    const repeat = ['daily', 'weekly', 'monthly'].includes(item.repeat) ? item.repeat : ''
    add('行事曆', item, 'event', { title: item.title, project: calendarNames.get(item.calId),
      body: [text(item.location), text(item.notes)].filter(Boolean).join('\n'), ...start,
      time: item.allDay ? '' : start.time, endDate: item.allDay && end.due ? addDays(end.due, -1) : end.due, endTime: item.allDay ? '' : end.time, repeat })
  }
  const contactState = load('macos27:app:contacts')
  for (const item of Array.isArray(contactState.contacts) ? contactState.contacts : []) {
    if (!item || (!item.first && !item.last)) continue
    add('通訊錄', item, 'contact', {
      title: [text(item.first), text(item.last)].filter(Boolean).join(' '), email: text(item.email, 320),
      phone: text(item.phone, 80), company: text(item.company, 250),
      body: [text(item.address), text(item.url), text(item.birthday), text(item.note)].filter(Boolean).join('\n'),
    })
  }
  const fs = load('macos27:fs')
  for (const item of Object.values(fs.nodes ?? {})) {
    if (!item || item.trashed || item.kind !== 'file' || typeof item.content !== 'string' || item.content.startsWith('data:')) continue
    if (item.mime?.startsWith('text/') || /\.(txt|md|log)$/i.test(item.name ?? '')) {
      add('Finder', item, 'doc', { title: item.name, body: text(item.content, 2_000_000) })
    }
  }
  return entries
}
