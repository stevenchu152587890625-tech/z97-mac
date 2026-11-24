const webEndpoint = new URL('../api/search', import.meta.url)
const sources = { web: '網頁', wiki: '維基百科' }
const providerNames = { 'bing-rss': 'Bing RSS', wikipedia: '維基百科' }

export function searchQueryFromURL(url) {
  if (!url || url.pathname.replace(/\/$/, '') !== '/search') return ''
  const host = url.hostname.toLowerCase().replace(/^www\./, '')
  if (!['google.com', 'google.com.tw', 'bing.com'].includes(host)) return ''
  return url.searchParams.get('q')?.trim() ?? ''
}

function hrefFor(address) {
  try {
    const url = new URL(address)
    return ['http:', 'https:'].includes(url.protocol) && !url.username && !url.password ? url : null
  } catch { return null }
}

function wikipediaSnippet(value) {
  const text = (typeof value === 'string' ? value : '').replace(/<[^>]*>/g, '')
  const template = document.createElement('template')
  template.innerHTML = text.replace(/</g, '&lt;')
  return template.content.textContent.replace(/\s+/g, ' ').trim().slice(0, 450)
}

async function requestJSON(url, options, signal, timeout) {
  const deadline = AbortSignal.timeout(timeout)
  const response = await fetch(url, {
    credentials: 'omit', referrerPolicy: 'no-referrer', ...options,
    signal: AbortSignal.any([signal, deadline]),
  })
  if (!response.ok || !response.headers.get('Content-Type')?.includes('application/json')) {
    throw new Error('搜尋服務暫時無法回應。')
  }
  return response.json()
}

async function wikiSearch(query, signal) {
  const language = /\p{Script=Han}/u.test(query) ? 'zh' : 'en'
  const api = new URL('https://' + language + '.wikipedia.org/w/api.php')
  api.search = new URLSearchParams({
    action: 'query', list: 'search', srsearch: query, srnamespace: '0',
    srlimit: '12', format: 'json', origin: '*',
  })
  const data = await requestJSON(api, {}, signal, 6500)
  if (!Array.isArray(data.query?.search)) throw new Error('搜尋服務未回傳結果清單。')
  return data.query.search.filter(item => typeof item.title === 'string' && Number.isSafeInteger(item.pageid))
    .map(item => ({
      title: item.title,
      url: 'https://' + language + '.wikipedia.org/?curid=' + item.pageid,
      snippet: wikipediaSnippet(item.snippet),
    }))
}

async function generalSearch(query, signal) {
  const data = await requestJSON(webEndpoint, {
    method: 'POST', credentials: 'same-origin', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ q: query }),
  }, signal, 8500)
  if (data.source !== 'bing-rss' || !Array.isArray(data.results)) throw new Error('搜尋服務未回傳結果清單。')
  return data.results
}

export function webSearchView(jsx, react, { query, navigate, local, online }) {
  const { jsx: h, jsxs: hs } = jsx
  const [scope, setScope] = react.useState('web')
  const [revision, retry] = react.useState(0)
  const [state, setState] = react.useState(null)
  const normalized = query.trim()
  const tooLong = Array.from(normalized).length > 200

  react.useEffect(() => {
    if (!normalized || tooLong || !online) return
    const controller = new AbortController()
    const { signal } = controller
    const report = update => {
      if (!signal.aborted) setState({ query: normalized, scope, revision, ...update })
    }
    report({ phase: 'loading' })
    const run = async () => {
      if (scope === 'web') {
        try {
          const results = await generalSearch(normalized, signal)
          report({ phase: 'done', results, provider: 'bing-rss' })
          return
        } catch {
          if (signal.aborted) return
          report({ phase: 'loading', fallback: true })
        }
      }
      try {
        const results = await wikiSearch(normalized, signal)
        report({ phase: 'done', results, provider: 'wikipedia', fallback: scope === 'web' })
      } catch {
        report({ phase: 'error' })
      }
    }
    run()
    return () => controller.abort()
  }, [normalized, scope, revision, tooLong, online])

  const current = state?.query === normalized && state.scope === scope && state.revision === revision ? state : null
  const results = (online && !tooLong && current?.results ? current.results : [])
    .map(item => ({ ...item, address: hrefFor(item.url) }))
    .filter(item => item.address && typeof item.title === 'string' && item.title.trim())
  const message = !online ? 'Wi-Fi 已關閉。仍可查看下方的本機結果。'
    : tooLong ? '查詢最多 200 個字元。'
    : !normalized ? '請在 Safari 網址列輸入關鍵字。'
    : current?.phase === 'error' ? '搜尋服務暫時無法使用。'
    : current?.phase === 'done' ? (results.length ? results.length + ' 筆結果' : '沒有找到符合的結果。')
    : current?.fallback ? '網頁搜尋暫無回應，正在查詢維基百科…' : '搜尋中…'

  return h('div', { className: 'z97-search', children: hs('section', {
    className: 'z97-search-body',
    children: [
      h('h1', { children: normalized || '搜尋' }),
      hs('div', { className: 'z97-search-toolbar', children: [
        ...Object.entries(sources).map(([key, label]) => h('button', {
          type: 'button', 'aria-pressed': scope === key, onClick: () => setScope(key), children: label,
        }, key)),
        current?.phase === 'error' ? h('button', { type: 'button', onClick: () => retry(value => value + 1), children: '重新查詢' }) : null,
      ] }),
      hs('p', { className: 'z97-search-status', role: 'status', children: [
        message,
        online && current?.phase === 'done' ? ' · 來源 ' + providerNames[current.provider] : '',
        online && current?.phase === 'done' && current.fallback ? '（網頁搜尋暫無回應）' : '',
      ] }),
      results.length ? h('ol', { className: 'z97-search-list', children: results.map(item => hs('li', {
        children: [
          h('span', { className: 'z97-search-domain', children: item.address.hostname }),
          h('a', { href: item.address.href, target: '_blank', rel: 'noopener noreferrer', children: item.title }),
          item.snippet ? h('p', { children: item.snippet }) : null,
        ],
      }, item.address.href)) }) : null,
      local.length ? hs('section', { className: 'z97-search-local', children: [
        h('h2', { children: '書籤與內建網站' }),
        h('ul', { children: local.map(item => h('li', {
          children: h('button', { type: 'button', onClick: () => navigate(item.url), children: item.title }),
        }, item.url)) }),
      ] }) : null,
    ],
  }) })
}
