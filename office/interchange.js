import { addDays } from './calendar.js'
import { localDate } from './formats.js'

const localTime = date => [date.getHours(), date.getMinutes()].map(number => String(number).padStart(2, '0')).join(':')
const unescape = value => (value ?? '').replace(/\\([\\;,nN])/g, (_, character) => /n/i.test(character) ? '\n' : character)

function property(line) {
  let quoted = false
  for (let i = 0; i < line.length; i++) {
    if (line[i] === '"') quoted = !quoted
    if (line[i] !== ':' || quoted) continue
    const head = line.slice(0, i)
    const name = head.split(';')[0].split('.').pop().toUpperCase()
    return { name, head, value: line.slice(i + 1) }
  }
  return null
}

function items(text, component) {
  const lines = text.replace(/^\uFEFF/, '').replace(/(?:\r\n|\n|\r)[ \t]/g, '').split(/\r\n|\n|\r/)
  const output = []
  let fields
  let level = 0
  for (const line of lines) {
    const item = property(line)
    if (!item) continue
    if (item.name === 'BEGIN' && item.value.toUpperCase() === component && !fields) {
      fields = new Map()
      level = 0
      continue
    }
    if (!fields) continue
    if (item.name === 'BEGIN') level++
    else if (item.name === 'END') {
      if (level) level--
      else if (item.value.toUpperCase() === component) {
        output.push(fields)
        fields = undefined
        if (output.length > 1000) throw new Error('單個匯入檔最多支援 1,000 筆資料。')
      }
    } else if (level === 0 && !fields.has(item.name)) fields.set(item.name, item)
  }
  return output
}

function splitEscaped(value, separator) {
  const pieces = []
  let part = ''
  let escaped = false
  for (const character of value ?? '') {
    if (character === separator && !escaped) { pieces.push(unescape(part)); part = '' }
    else part += character
    if (character === '\\') escaped = !escaped
    else escaped = false
  }
  pieces.push(unescape(part))
  return pieces
}

export function readContacts(text) {
  const cards = items(text, 'VCARD')
  const entries = []
  let skipped = 0
  for (const card of cards) {
    const data = key => card.get(key)?.value ?? ''
    if (!['3.0', '4.0'].includes(data('VERSION').trim())) { skipped++; continue }
    const name = unescape(data('FN')) || splitEscaped(data('N'), ';').filter(Boolean).join(' ')
    if (!name.trim()) { skipped++; continue }
    entries.push({
      title: name.trim().slice(0, 250),
      company: splitEscaped(data('ORG'), ';').filter(Boolean).join(' / ').slice(0, 250),
      phone: unescape(data('TEL')).replace(/^tel:/i, '').slice(0, 80),
      email: unescape(data('EMAIL')).replace(/^mailto:/i, '').slice(0, 320),
      body: [unescape(data('NOTE')), splitEscaped(data('ADR'), ';').filter(Boolean).join(' ')]
        .filter(Boolean).join('\n').slice(0, 100000),
    })
  }
  if (!entries.length) throw new Error('找不到可匯入的聯絡人。工作台支援 UTF-8 vCard 3.0／4.0。')
  return { entries, skipped }
}

function parseMoment(field) {
  if (!field) return null
  const match = /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?$/.exec(field.value)
  if (!match) return null
  const [, year, month, day, hours, minutes, seconds = '00', utc] = match
  const date = year + '-' + month + '-' + day
  const iso = date + 'T' + (hours ?? '00') + ':' + (minutes ?? '00') + ':' + seconds
  const expected = Date.parse(iso + 'Z')
  if (!Number.isFinite(expected) || new Date(expected).toISOString().slice(0, 19) !== iso) return null
  if (!hours) return { date, time: '' }
  const zone = /(?:^|;)TZID=(?:"([^"]+)"|([^;]+))/i.exec(field.head ?? '')
  if (!utc && !zone) return { date, time: hours + ':' + minutes }
  let timestamp = expected
  if (!utc) {
    try {
      const formatter = new Intl.DateTimeFormat('en-GB', { timeZone: zone[1] || zone[2],
        year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
      const displayTimestamp = timestamp => {
        const entries = formatter.formatToParts(new Date(timestamp)).map(({ type, value }) => [type, value])
        const { year, month, day, hour, minute, second } = Object.fromEntries(entries)
        return Date.parse(year + '-' + month + '-' + day + 'T' + hour + ':' + minute + ':' + second + 'Z')
      }
      for (let i = 0; i < 4; i++) timestamp += expected - displayTimestamp(timestamp)
      if (displayTimestamp(timestamp) !== expected) return null
    } catch { return null }
  }
  const value = new Date(timestamp)
  const sourceZone = utc ? 'UTC' : (zone[1] || zone[2])
  const sourceClock = new Intl.DateTimeFormat('en', { timeZone: sourceZone }).resolvedOptions().timeZone
  const localClock = new Intl.DateTimeFormat('en').resolvedOptions().timeZone
  return { date: localDate(value), time: localTime(value), zoned: sourceClock !== localClock }
}

function recurrence(event, start) {
  const source = event.get('RRULE')?.value
  if (!source) return { repeat: '', reduced: event.has('RDATE') }
  const options = new Map(source.toUpperCase().split(';').map(value => value.split('=')))
  const repeat = options.get('FREQ')?.toLowerCase()
  if (event.has('EXDATE') || event.has('RDATE') || start.zoned ||
      !['daily', 'weekly', 'monthly'].includes(repeat) ||
      [...options.keys()].some(key => !['FREQ', 'UNTIL', 'INTERVAL'].includes(key)) ||
      (options.has('INTERVAL') && options.get('INTERVAL') !== '1')) return { repeat: '', reduced: true }
  const until = parseMoment({ value: options.get('UNTIL') ?? '' })
  return { repeat, repeatUntil: until?.date ?? '' }
}

export function readCalendar(text) {
  const entries = []
  let skipped = 0
  let simplified = 0
  for (const event of items(text, 'VEVENT')) {
    const value = key => event.get(key)?.value ?? ''
    const start = parseMoment(event.get('DTSTART'))
    const end = parseMoment(event.get('DTEND'))
    if (!start || value('STATUS').toUpperCase() === 'CANCELLED' || event.has('RECURRENCE-ID')) {
      skipped++
      continue
    }
    const rule = recurrence(event, start)
    const limited = rule.reduced || event.has('DURATION') || (event.has('DTEND') && !end)
    if (limited) simplified++
    entries.push({
      title: unescape(value('SUMMARY')).trim().slice(0, 250) || '未命名事件',
      due: start.date, time: start.time, repeat: rule.repeat, repeatUntil: rule.repeatUntil,
      endDate: !start.time && end?.date ? addDays(end.date, -1) : end?.date ?? '',
      endTime: start.time ? end?.time ?? '' : '',
      body: [unescape(value('LOCATION')), unescape(value('DESCRIPTION'))].filter(Boolean).join('\n').slice(0, 100000),
    })
  }
  if (!entries.length) throw new Error('沒有可匯入的 ICS 事件。請確認檔案格式與時區。')
  return { entries, skipped, simplified }
}
