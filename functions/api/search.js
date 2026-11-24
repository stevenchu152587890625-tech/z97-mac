const MAX_BYTES = 512 * 1024

function reply(status, body) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' },
  })
}

async function boundedText(body, limit) {
  if (!body) throw new Error('Missing response body')
  const reader = body.getReader()
  const decoder = new TextDecoder()
  let size = 0
  let text = ''
  try {
    while (true) {
      const { done, value } = await reader.read()
      if (done) return text + decoder.decode()
      size += value.byteLength
      if (size > limit) throw new Error('Response too large')
      text += decoder.decode(value, { stream: true })
    }
  } finally {
    await reader.cancel().catch(() => {})
  }
}

function decode(text) {
  const codes = { amp: '&', quot: '"', apos: "'", lt: '<', gt: '>', nbsp: ' ' }
  return text.replace(/&(#x[0-9a-f]+|#\d+|amp|quot|apos|lt|gt|nbsp);/gi, (whole, entity) => {
    if (entity[0] !== '#') return codes[entity.toLowerCase()] ?? whole
    const hex = entity[1].toLowerCase() === 'x'
    const value = Number.parseInt(entity.slice(hex ? 2 : 1), hex ? 16 : 10)
    return value > 0 && value <= 0x10ffff ? String.fromCodePoint(value) : ''
  })
}

function field(item, name) {
  const value = new RegExp('<' + name + '(?:\\s[^>]*)?>([\\s\\S]*?)</' + name + '>', 'i').exec(item)?.[1] ?? ''
  return decode(value.replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')).trim()
}

export function readFeed(xml) {
  if (!/<rss(?:\s|>)/i.test(xml) || !/<channel(?:\s|>)/i.test(xml) || /<!DOCTYPE|<!ENTITY/i.test(xml)) {
    throw new Error('Unusable search feed')
  }
  const results = []
  const seen = new Set()
  const clean = (text, limit) => decode(text.replace(/<[^>]*>/g, ' ')).replace(/\s+/g, ' ').trim().slice(0, limit)
  for (const match of xml.matchAll(/<item(?:\s[^>]*)?>([\s\S]*?)<\/item>/gi)) {
    const item = match[1]
    const address = field(item, 'link')
    const title = clean(field(item, 'title'), 300)
    if (!title || address.length > 2500) continue
    let url
    try { url = new URL(address) } catch { continue }
    if (!['http:', 'https:'].includes(url.protocol) || url.username || url.password || seen.has(url.href)) continue
    seen.add(url.href)
    results.push({ url: url.href, title, snippet: clean(field(item, 'description'), 500) })
    if (results.length === 12) break
  }
  return results
}

export async function onRequestPost({ request }) {
  let query
  try {
    if (!/^application\/json(?:\s*;|$)/i.test(request.headers.get('content-type') ?? '')) throw new Error('Invalid content type')
    const data = JSON.parse(await boundedText(request.body, 4096))
    query = typeof data?.q === 'string' ? data.q.trim() : ''
    if (!query || Array.from(query).length > 200) throw new Error('Invalid query')
  } catch {
    return reply(400, { error: 'invalid_query' })
  }

  try {
    const endpoint = new URL('https://www.bing.com/search')
    endpoint.search = new URLSearchParams({ q: query, format: 'rss', mkt: 'zh-TW', setlang: 'zh-TW' })
    const response = await fetch(endpoint, {
      headers: { Accept: 'application/rss+xml, text/xml' },
      signal: AbortSignal.timeout(7000),
    })
    if (!response.ok) throw new Error('Upstream returned ' + response.status)
    return reply(200, { source: 'bing-rss', results: readFeed(await boundedText(response.body, MAX_BYTES)) })
  } catch {
    return reply(502, { error: 'provider_unavailable' })
  }
}
