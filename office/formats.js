import { eventEnd } from './calendar.js'
export function element(tag, attrs = {}, ...children) {
  const node = document.createElement(tag)
  for (const [key, value] of Object.entries(attrs)) {
    if (value === undefined || value === null || value === false) continue
    if (key.startsWith('on')) node.addEventListener(key.slice(2), value)
    else if (key === 'class') node.className = value
    else node.setAttribute(key, value === true ? '' : value)
  }
  for (const child of children.flat()) {
    if (child !== undefined && child !== null) node.append(child)
  }
  return node
}

export function localDate(date = new Date()) {
  const pad = n => String(n).padStart(2, '0')
  return date.getFullYear() + '-' + pad(date.getMonth() + 1) + '-' + pad(date.getDate())
}

export function download(filename, data, type) {
  const url = URL.createObjectURL(data instanceof Blob ? data : new Blob([data], { type }))
  const link = element('a', { href: url, download: filename })
  document.body.append(link)
  link.click()
  link.remove()
  setTimeout(() => URL.revokeObjectURL(url), 2000)
}

export function filename(title) {
  return title.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').replace(/[ .]+$/, '').slice(0, 150) || '未命名'
}

function inline(text) {
  const markup = /(\x60[^\x60]+\x60|\*\*[^*]+\*\*|\*[^*]+\*|\[[^\]]+\]\([^\s)]+\))/g
  const nodes = []
  let end = 0
  for (const match of text.matchAll(markup)) {
    nodes.push(text.slice(end, match.index))
    const token = match[0]
    if (token.startsWith('**')) nodes.push(element('strong', {}, token.slice(2, -2)))
    else if (token.startsWith('*')) nodes.push(element('em', {}, token.slice(1, -1)))
    else if (token.startsWith('[')) {
      const [, label, url] = token.match(/^\[([^\]]+)\]\((.+)\)$/)
      let safe = false
      try { safe = ['https:', 'http:', 'mailto:'].includes(new URL(url).protocol) } catch {}
      nodes.push(safe ? element('a', { href: url, target: '_blank', rel: 'noopener noreferrer' }, label) : label)
    } else nodes.push(element('code', {}, token.slice(1, -1)))
    end = match.index + token.length
  }
  nodes.push(text.slice(end))
  return nodes
}

export function markdown(text) {
  const preview = element('div', { class: 'markdown' })
  let list
  let code
  for (const line of text.split(/\r?\n/)) {
    if (/^\s*\x60{3}/.test(line)) {
      if (code) {
        preview.append(element('pre', {}, code.join('\n')))
        code = undefined
      } else {
        code = []
        list = undefined
      }
      continue
    }
    if (code) {
      code.push(line)
      continue
    }
    const item = line.match(/^\s*(?:([-*])|(\d+)\.)\s+(.+)/)
    if (item) {
      const kind = item[2] ? 'ol' : 'ul'
      if (list?.localName !== kind) {
        list = element(kind, item[2] ? { start: item[2] } : {})
        preview.append(list)
      }
      list.append(element('li', {}, inline(item[3])))
      continue
    }
    list = undefined
    const heading = line.match(/^(#{1,6})\s+(.+)/)
    if (heading) preview.append(element('h' + heading[1].length, {}, inline(heading[2])))
    else if (line.startsWith('> ')) preview.append(element('blockquote', {}, inline(line.slice(2))))
    else if (/^\s*---+\s*$/.test(line)) preview.append(element('hr'))
    else if (line.trim()) preview.append(element('p', {}, inline(line)))
  }
  if (code) preview.append(element('pre', {}, code.join('\n')))
  return preview
}

const icsEscape = text => String(text ?? '').replace(/\\/g, '\\\\').replace(/\r\n?|\n/g, '\\n').replace(/[;,]/g, '\\$&')
function fold(line) {
  const encoder = new TextEncoder()
  let bytes = 0
  let result = ''
  for (const char of line) {
    const size = encoder.encode(char).length
    if (bytes + size > 75) {
      result += '\r\n '
      bytes = 1
    }
    result += char
    bytes += size
  }
  return result
}

export function writeCalendar(events) {
  const stamp = new Date().toISOString().replace(/[-:]|\.\d+/g, '')
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Z97//Office//ZH', 'CALSCALE:GREGORIAN']
  for (const event of events) {
    if (!event.due) continue
    const date = event.due.replaceAll('-', '')
    const time = event.time?.replace(':', '')
    lines.push('BEGIN:VEVENT', 'UID:' + event.id + '@z97-office', 'DTSTAMP:' + stamp,
      time ? 'DTSTART:' + date + 'T' + time + '00' : 'DTSTART;VALUE=DATE:' + date,
      'SUMMARY:' + icsEscape(event.title), 'DESCRIPTION:' + icsEscape(event.body))
    const end = eventEnd(event)
    if (end) {
      const endDate = end.date.replaceAll('-', '')
      lines.push(end.time ? 'DTEND:' + endDate + 'T' + end.time.replace(':', '') + '00' : 'DTEND;VALUE=DATE:' + endDate)
    }
    if (['daily', 'weekly', 'monthly'].includes(event.repeat)) {
      const limit = event.repeatUntil?.replaceAll('-', '')
      lines.push('RRULE:FREQ=' + event.repeat.toUpperCase() + (limit ? ';UNTIL=' + limit + (time ? 'T235959' : '') : ''))
    }
    lines.push('END:VEVENT')
  }
  lines.push('END:VCALENDAR')
  return lines.map(fold).join('\r\n') + '\r\n'
}
