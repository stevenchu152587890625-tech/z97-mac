const byApp = { notes: 'doc', textedit: 'doc', reminders: 'task', calendar: 'event', contacts: 'contact' }
const finderViews = { home: '工作總覽', sheet: '試算表', asset: '檔案庫', trash: '工作資料垃圾桶', settings: '備份與資料移轉' }

function viewFor(appId, payload) {
  if (appId === 'finder') return Object.hasOwn(finderViews, payload?.officeView) ? payload.officeView : ''
  if (!Object.hasOwn(byApp, appId)) return ''
  if (payload && Object.keys(payload).some(key => key !== '_focusOnly')) return ''
  return byApp[appId]
}

export function desktopTitle(appId, payload, original) {
  return appId === 'finder' && viewFor(appId, payload) ? finderViews[payload.officeView] : original
}

function frameFor(windowId) {
  return windowId && document.querySelector('iframe[data-office-window="' + CSS.escape(windowId) + '"]')
}

function bridgeFor(windowId) {
  return frameFor(windowId)?.contentWindow?.z97Desktop
}

export function officeContent(jsx, Original, appId, windowId, payload, focus, props) {
  const view = viewFor(appId, payload)
  if (!view) return jsx.jsx(Original, props)
  const url = new URL('./', import.meta.url)
  url.searchParams.set('app', appId)
  url.hash = view
  return jsx.jsx('iframe', {
    'data-office-window': windowId,
    title: appId === 'finder' ? finderViews[view] : '辦公功能',
    src: url.href,
    style: { display: 'block', width: '100%', height: '100%', border: 0, background: '#f6f7f9' },
    onLoad: event => {
      const doc = event.currentTarget.contentDocument
      doc?.addEventListener('pointerdown', focus, true)
      doc?.addEventListener('focusin', focus, true)
    },
  })
}

const appMenus = {
  notes: { label: 'Note', id: 'note', name: '備忘錄' },
  textedit: { label: 'File', id: 'file', name: '文字編輯' },
  reminders: { label: 'Reminder', id: 'reminder', name: '提醒事項' },
  calendar: { label: 'Event', id: 'cal-event', name: '行事曆' },
  contacts: { label: 'Contact', id: 'contact', name: '聯絡人' },
}
const typeByView = { doc: '文件', sheet: '試算表', task: '任務', event: '事件', contact: '聯絡人' }
const separator = { separator: true }

export function desktopMenus(appId, windowId, payload, launch, close, original = []) {
  const view = viewFor(appId, payload)
  const command = (label, name, shortcut) => ({ label, shortcut, action: () => bridgeFor(windowId)?.command(name) })
  const closeWindow = { label: 'Close Window', shortcut: '⌘W', disabled: !windowId, action: close }
  if (appId === 'finder') {
    const views = Object.entries(finderViews).map(([officeView, label]) => ({
      label, action: () => launch('finder', { officeView }),
    }))
    if (view) {
      const items = [command('搜尋…', 'search', '⌘F')]
      if (view === 'sheet') items.unshift(command('新增試算表', 'new', '⌘N'), command('匯入試算表…', 'import', '⌘O'))
      if (view === 'asset') items.unshift(command('匯入檔案…', 'import', '⌘O'))
      if (view === 'settings') items.unshift(command('下載完整備份', 'backup'), command('還原備份…', 'restore'))
      return [{ id: 'file', label: 'File', items: [...items, separator, closeWindow] }, { id: 'go', label: 'Go', items: views }]
    }
    let foundGo = false
    const menus = original.map(menu => {
      if (menu.id !== 'go' && menu.label !== 'Go') return menu
      foundGo = true
      return { ...menu, items: [...menu.items, separator, ...views] }
    })
    return foundGo ? menus : [...menus, { id: 'go', label: 'Go', items: views }]
  }
  if (!view) return original
  const { id, label, name } = appMenus[appId]
  const items = [
    command('新增' + typeByView[view], 'new', '⌘N'),
    command('搜尋…', 'search', '⌘F'),
    command('儲存', 'save', '⌘S'),
  ]
  if (['doc', 'event', 'contact'].includes(view)) items.push(command('匯入…', 'import', '⌘O'))
  if (view === 'doc') items.push(command('下載 Word', 'word'), command('下載 Markdown', 'markdown', '⇧⌘E'), command('列印／PDF…', 'print', '⌘P'))
  if (view === 'event') items.push(command('今天', 'today', '⌘T'), command('匯出 ICS', 'ics'))
  if (view === 'contact') items.push(command('下載 vCard', 'vcard'))
  items.push(separator, { label: '開啟原版' + name, action: () => launch(appId, { officeLegacy: true }) })
  if (id === 'file') return [{ id, label, items: [...items, separator, closeWindow] }]
  const file = view === 'doc'
    ? [items[0], ...items.slice(2).filter(item => ['⌘S', '⌘O', '⌘P'].includes(item.shortcut))]
    : [items[0]]
  return [{ id, label, items }, { id: 'file', label: 'File', items: [...file, separator, closeWindow] }]
}

export function saveBeforeClosing(ids, close) {
  const pending = ids.map(id => bridgeFor(id)).filter(bridge => bridge?.hasPending())
  if (!pending.length) return close()
  Promise.all(pending.map(bridge => bridge.flush())).then(results => {
    if (results.every(Boolean)) close()
  }).catch(error => console.error('無法在關閉視窗前儲存工作資料', error))
}
