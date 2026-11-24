function ordinal(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value ?? '')) return NaN
  return Date.parse(value + 'T00:00:00Z') / 86_400_000
}

export function addDays(value, days) {
  const date = new Date((ordinal(value) + days) * 86_400_000)
  return Number.isFinite(date.getTime()) ? date.toISOString().slice(0, 10) : ''
}

export function eventEnd(event) {
  if (!event.due) return null
  const end = event.endDate || (event.endTime ? event.due : '')
  if (!end) return null
  if (end < event.due) return null
  if (!event.time) return { date: addDays(end, 1), time: '' }
  const time = event.endTime || event.time
  if (end === event.due && time <= event.time) return null
  return { date: end, time }
}

export function occursOn(event, day) {
  if (!event.due || day < event.due) return false
  if (event.kind !== 'event') return event.due === day
  const start = ordinal(event.due)
  const target = ordinal(day)
  const end = eventEnd(event)
  const duration = end ? Math.max(0, ordinal(end.date) - start - (end.time && end.time !== '00:00' ? 0 : 1)) : 0
  const finalStart = Math.min(target, event.repeatUntil ? ordinal(event.repeatUntil) : Infinity)
  const earliest = Math.max(target - duration, start)
  if (finalStart < earliest) return false
  if (event.repeat === 'daily') return true
  if (event.repeat === 'weekly') {
    const nearest = start + 7 * Math.floor((finalStart - start) / 7)
    return nearest >= earliest
  }
  if (event.repeat === 'monthly') {
    const date = new Date(finalStart * 86_400_000)
    const dayOfMonth = +event.due.slice(8, 10)
    for (let i = 0; i < 3; i++) {
      const candidate = new Date(date.getTime())
      candidate.setUTCDate(1)
      candidate.setUTCMonth(candidate.getUTCMonth() - i)
      candidate.setUTCDate(dayOfMonth)
      if (candidate.getUTCDate() !== dayOfMonth) continue
      const occurrence = candidate.getTime() / 86_400_000
      if (occurrence <= finalStart) return occurrence >= earliest
    }
    return false
  }
  return start >= earliest
}

export function nextOccurrence(event, today, horizon = 45) {
  const limit = addDays(today, horizon)
  if (event.due > limit) return event.due
  for (let day = today; day && day <= limit; day = addDays(day, 1)) {
    if (occursOn(event, day)) return day
  }
  return ''
}
