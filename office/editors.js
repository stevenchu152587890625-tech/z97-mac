import { excelType, wordType, writeExcel, writeWord } from './ooxml.js'
import { maxRows, maxColumns, statuses, types } from './storage.js'
import { columnName, evaluateSheet, parseDelimited, writeCSV } from './grid.js'
import { download, element as el, filename, markdown, writeCalendar } from './formats.js'

const priorities = { low: '低', normal: '一般', high: '高' }
export const button = (text, handler, className = 'secondary', disabled = false) =>
  el('button', { type: 'button', class: className, onclick: handler, disabled }, text)

function control(label, input) {
  return el('label', { class: 'field' }, el('span', {}, label), input)
}

function inputField(label, value, oninput, options = {}) {
  const input = el('input', { ...options, oninput: event => oninput(event.target.value) })
  input.value = value ?? ''
  return control(label, input)
}

function selectField(label, choices, value, handler) {
  const select = el('select', { onchange: event => handler(event.target.value) },
    Object.entries(choices).map(([key, name]) => el('option', { value: key }, name)))
  select.value = value ?? Object.keys(choices)[0]
  return control(label, select)
}

function description(label, value, handler) {
  const field = el('textarea', { class: 'details', rows: 6, maxlength: 100000, oninput: event => handler(event.target.value) })
  field.value = value ?? ''
  return control(label, field)
}

function documentEditor(record, editor) {
  const area = el('textarea', {
    class: 'document-text', 'aria-label': '編輯文件', maxlength: 2000000,
    placeholder: '開始撰寫，或從工具列插入格式…', spellcheck: 'false',
  })
  area.value = record.body ?? ''
  const preview = el('div', { class: 'document-preview', 'aria-label': '文件預覽' })
  const panes = el('div', { class: 'document-panes' }, area, preview)
  const wordCount = el('span', { class: 'muted document-count' })
  let previewFrame
  const redraw = () => {
    cancelAnimationFrame(previewFrame)
    previewFrame = requestAnimationFrame(() => {
      if (!panes.classList.contains('editor-only')) preview.replaceChildren(markdown(area.value))
      const words = area.value.match(/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}]|[\p{L}\p{N}]+/gu) ?? []
      wordCount.textContent = words.length.toLocaleString() + ' 字詞'
    })
  }
  area.addEventListener('input', () => { editor.change({ body: area.value }); redraw() })
  const insert = (prefix, suffix = '', placeholder = '') => {
    const { selectionStart: start, selectionEnd: end } = area
    const content = area.value.slice(start, end) || placeholder
    area.focus()
    area.setRangeText(prefix + content + suffix, start, end, 'end')
    area.setSelectionRange(start + prefix.length, start + prefix.length + content.length)
    editor.change({ body: area.value })
    redraw()
  }
  const mode = (variant, event) => {
    panes.className = 'document-panes ' + variant
    for (const node of viewButtons.children) node.setAttribute('aria-pressed', String(node === event.currentTarget))
    redraw()
  }
  const viewButtons = el('div', { class: 'segmented', 'aria-label': '文件顯示模式' },
    button('編輯', event => mode('editor-only', event), 'quiet'),
    button('並排', event => mode('', event), 'quiet'),
    button('預覽', event => mode('preview-only', event), 'quiet'))
  viewButtons.children[1].setAttribute('aria-pressed', 'true')
  redraw()
  return el('section', { class: 'panel document-panel' },
    el('div', { class: 'toolbar' },
      button('標題', () => insert('## ', '', '標題'), 'quiet'),
      button('粗體', () => insert('**', '**', '文字'), 'quiet'),
      button('清單', () => insert('- ', '', '項目'), 'quiet'),
      button('連結', () => insert('[', '](https://example.com)', '連結文字'), 'quiet'),
      viewButtons, wordCount),
    panes,
    el('div', { class: 'toolbar footer-tools' },
      button('下載 Word', () => download(filename(editor.draft.title) + '.docx', writeWord(editor.draft.title, area.value), wordType)),
      button('下載 Markdown', () => download(filename(editor.draft.title) + '.md', area.value, 'text/markdown;charset=utf-8')),
      button('下載純文字', () => download(filename(editor.draft.title) + '.txt', area.value, 'text/plain;charset=utf-8')),
      button('列印／PDF', () => {
        const paper = el('article', { class: 'print-document' }, el('h1', {}, editor.draft.title), markdown(area.value))
        document.body.append(paper)
        try { window.print() } finally { paper.remove() }
      })))
}

function sheetEditor(record, editor) {
  let rows = record.rows?.length ? record.rows.map(row => [...row]) : Array.from({ length: 12 }, () => Array(6).fill(''))
  let width = Math.max(1, ...rows.map(row => row.length))
  rows.forEach(row => { while (row.length < width) row.push('') })
  let selected = [0, 0]
  let firstRow = 0
  const pageSize = 40
  let visible = new Map()
  let calculationFrame
  const grid = el('div', { class: 'grid-scroll' })
  const range = el('span', { class: 'muted' })
  const ref = el('span', { class: 'cell-ref' }, 'A1')
  const formula = el('textarea', { rows: 1, 'aria-label': '儲存格內容或公式', maxlength: 20000,
    placeholder: '輸入內容或公式，例如 =SUM(A1:A5)' })
  const changed = () => editor.change({ rows })
  const keyFor = (r, c) => r + ':' + c
  const highlight = () => {
    ref.textContent = columnName(selected[1]) + (selected[0] + 1)
    range.textContent = (firstRow + 1) + '–' + Math.min(firstRow + pageSize, rows.length) + ' / ' + rows.length + ' 列'
    formula.value = rows[selected[0]]?.[selected[1]] ?? ''
    for (const [key, input] of visible) input.classList.toggle('selected-cell', key === keyFor(...selected))
  }
  const recalculate = () => {
    const calculate = evaluateSheet(rows)
    for (const [key, input] of visible) {
      if (input === document.activeElement) continue
      const [r, c] = key.split(':').map(Number)
      input.value = String(calculate(r, c))
      input.classList.toggle('formula-cell', rows[r][c].startsWith('='))
    }
  }
  const focusCell = (r, c) => {
    selected = [Math.max(0, Math.min(rows.length - 1, r)), Math.max(0, Math.min(width - 1, c))]
    if (selected[0] < firstRow || selected[0] >= firstRow + pageSize) {
      firstRow = Math.floor(selected[0] / pageSize) * pageSize
      draw()
    }
    highlight()
    const cell = visible.get(keyFor(...selected))
    cell?.focus()
    cell?.select()
  }
  const applyPaste = (event, r, c) => {
    const clipboard = event.clipboardData?.getData('text/plain').replace(/\r\n/g, '\n') ?? ''
    if (!/[\t\n]/.test(clipboard)) return
    event.preventDefault()
    try {
      const cells = parseDelimited(clipboard, '\t')
      const newWidth = Math.max(width, c + Math.max(...cells.map(row => row.length)))
      const newHeight = Math.max(rows.length, r + cells.length)
      if (newWidth > maxColumns || newHeight > maxRows) throw new Error('貼上範圍超過工作表上限。')
      rows = Array.from({ length: newHeight }, (_, row) =>
        Array.from({ length: newWidth }, (_, col) => cells[row - r]?.[col - c] ?? rows[row]?.[col] ?? ''))
      width = newWidth
      changed()
      draw()
      focusCell(r, c)
    } catch (error) {
      editor.notify(error.message)
    }
  }
  const draw = () => {
    visible = new Map()
    const header = el('tr', {}, el('th', { scope: 'col', 'aria-label': '列號' }),
      Array.from({ length: width }, (_, c) => el('th', { scope: 'col' }, columnName(c))))
    const body = el('tbody')
    for (let r = firstRow; r < Math.min(firstRow + pageSize, rows.length); r++) {
      const line = el('tr', {}, el('th', { scope: 'row' }, String(r + 1)))
      for (let c = 0; c < width; c++) {
        const label = columnName(c) + (r + 1)
        const cell = el('input', {
          type: 'text', 'aria-label': label, maxlength: 20000, spellcheck: 'false',
          onfocus: () => { selected = [r, c]; highlight(); cell.value = rows[r][c] },
          oninput: () => {
            selected = [r, c]
            rows[r][c] = cell.value
            highlight()
            changed()
            cancelAnimationFrame(calculationFrame)
            calculationFrame = requestAnimationFrame(recalculate)
          },
          onblur: () => { queueMicrotask(recalculate) },
          onpaste: event => applyPaste(event, r, c),
          onkeydown: event => {
            if (event.key === 'Enter') {
              event.preventDefault()
              focusCell(r + (event.shiftKey ? -1 : 1), c)
            }
            if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
              event.preventDefault()
              focusCell(r + (event.key === 'ArrowDown' ? 1 : -1), c)
            }
          },
        })
        visible.set(keyFor(r, c), cell)
        line.append(el('td', {}, cell))
      }
      body.append(line)
    }
    grid.replaceChildren(el('table', { class: 'sheet-grid', 'aria-label': '工作表儲存格' }, el('thead', {}, header), body))
    highlight()
    recalculate()
  }
  const page = delta => {
    const target = Math.min(Math.floor((rows.length - 1) / pageSize) * pageSize, Math.max(0, firstRow + delta * pageSize))
    firstRow = target
    selected[0] = target
    draw()
  }
  const add = axis => {
    if ((axis === 'r' && rows.length >= maxRows) || (axis === 'c' && width >= maxColumns)) {
      editor.notify('已達工作表列欄上限。')
      return
    }
    if (axis === 'r') {
      rows.push(Array(width).fill(''))
      firstRow = Math.floor((rows.length - 1) / pageSize) * pageSize
      selected[0] = rows.length - 1
    } else {
      width++
      rows.forEach(row => row.push(''))
      selected[1] = width - 1
    }
    changed()
    draw()
    focusCell(...selected)
  }
  formula.addEventListener('input', () => {
    rows[selected[0]][selected[1]] = formula.value
    changed()
    recalculate()
  })
  formula.addEventListener('keydown', event => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault()
      focusCell(...selected)
    }
  })
  draw()
  return el('section', { class: 'panel sheet-panel' },
    el('div', { class: 'toolbar' },
      button('新增列', () => add('r')), button('新增欄', () => add('c')),
      button('清除儲存格', () => {
        rows[selected[0]][selected[1]] = ''
        changed()
        draw()
        focusCell(...selected)
      }),
      button('下載 Excel', () => download(filename(editor.draft.title) + '.xlsx', writeExcel(rows), excelType)),
      button('下載 CSV', () => download(filename(editor.draft.title) + '.csv', writeCSV(rows), 'text/csv;charset=utf-8'))),
    el('div', { class: 'formula-bar' }, ref, formula),
    grid,
    el('div', { class: 'toolbar footer-tools' },
      button('上一頁', () => page(-1), 'quiet'), range, button('下一頁', () => page(1), 'quiet'),
      el('span', { class: 'muted' }, '可貼上多格資料。支援 SUM、AVERAGE、MIN、MAX、COUNT、ROUND。以單引號開頭可保留文字。')))
}

function datedEditor(record, editor) {
  const change = key => value => editor.change({ [key]: value })
  const fields = [
    inputField(record.kind === 'task' ? '截止日' : '日期', record.due, change('due'), { type: 'date' }),
    inputField('時間', record.time, change('time'), { type: 'time' }),
  ]
  if (record.kind === 'event') fields.push(
    inputField('結束日期', record.endDate, change('endDate'), { type: 'date' }),
    inputField('結束時間', record.endTime, change('endTime'), { type: 'time' }),
    selectField('重複', { '': '不重複', daily: '每天', weekly: '每週', monthly: '每月' }, record.repeat, change('repeat')),
    inputField('重複至', record.repeatUntil, change('repeatUntil'), { type: 'date' }))
  if (record.kind === 'task') fields.push(
    selectField('狀態', statuses, record.status, change('status')),
    selectField('優先順序', priorities, record.priority ?? 'normal', change('priority')))
  return el('section', { class: 'panel form-panel' }, el('div', { class: 'form-grid' }, fields),
    description('備註', record.body, change('body')),
    button('下載行事曆 ICS', () => {
      const draft = editor.draft
      if (!draft.due) return editor.notify('請先選擇日期。')
      if (draft.endDate && draft.endDate < draft.due) return editor.notify('結束日期須在開始日期當天或之後。')
      if (draft.endTime && !draft.time) return editor.notify('請一併填寫開始時間。')
      if (draft.time && draft.endTime && (!draft.endDate || draft.endDate === draft.due) && draft.endTime <= draft.time) {
        return editor.notify('結束時間須晚於開始時間。')
      }
      if (draft.repeat && draft.repeatUntil && draft.repeatUntil < draft.due) return editor.notify('重複截止日須在開始後。')
      download(filename(editor.draft.title) + '.ics', writeCalendar([editor.draft]), 'text/calendar;charset=utf-8')
    }))
}

function contactEditor(record, editor) {
  const change = key => value => editor.change({ [key]: value })
  const email = () => (editor.draft.email ?? '').trim()
  return el('section', { class: 'panel form-panel' },
    el('div', { class: 'form-grid' },
      inputField('Email', record.email, change('email'), { type: 'email', maxlength: 320, autocomplete: 'off' }),
      inputField('電話', record.phone, change('phone'), { type: 'tel', maxlength: 80, autocomplete: 'off' }),
      inputField('組織', record.company, change('company'), { maxlength: 250 })),
    description('備註', record.body, change('body')),
    el('div', { class: 'toolbar' },
      button('用郵件程式撰寫', () => {
        if (!/^[^\s@]+@[^\s@]+$/.test(email())) return editor.notify('請先填寫有效的 Email。')
        location.href = 'mailto:' + encodeURIComponent(email()).replace('%40', '@')
      }),
      button('下載 vCard', () => {
        const escape = text => String(text ?? '').replace(/\\/g, '\\\\').replace(/\r\n?|\n/g, '\\n').replace(/[;,]/g, '\\$&')
        const draft = editor.draft
        const lines = ['BEGIN:VCARD', 'VERSION:3.0', 'N:' + escape(draft.title) + ';;;;', 'FN:' + escape(draft.title)]
        for (const [property, value] of [['ORG', draft.company], ['TEL', draft.phone], ['EMAIL', draft.email], ['NOTE', draft.body]]) {
          if (value) lines.push(property + ':' + escape(value))
        }
        download(filename(draft.title) + '.vcf', lines.concat('END:VCARD', '').join('\r\n'), 'text/vcard;charset=utf-8')
      })))
}

function assetEditor(record, editor) {
  const mime = (record.mime ?? '').toLowerCase()
  const file = record.blob
  const preview = el('div', { class: 'file-preview' })
  if (file && /^image\/(?:png|jpeg|gif|webp|avif|bmp)$/.test(mime)) {
    preview.append(el('img', { src: editor.objectURL(file), alt: record.title }))
  } else if (file && /^audio\/(?:mpeg|ogg|wav|mp4|flac|webm|x-wav)$/.test(mime)) {
    preview.append(el('audio', { src: editor.objectURL(file), controls: true, preload: 'metadata' }))
  } else if (file && /^video\/(?:mp4|webm|ogg|quicktime)$/.test(mime)) {
    preview.append(el('video', { src: editor.objectURL(file), controls: true, preload: 'metadata' }))
  } else if (file && /^text\/(?:plain|markdown|csv)$/.test(mime)) {
    file.slice(0, 300000).text().then(text => preview.append(el('pre', {}, text))).catch(error => editor.notify(error.message))
  } else {
    preview.append(el('p', { class: 'muted' }, '可下載並使用原有應用程式開啟。'))
  }
  return el('section', { class: 'panel form-panel' },
    el('p', { class: 'muted' }, record.filename + ' · ' + (file?.size ?? 0).toLocaleString() + ' bytes'),
    preview, button('下載原始檔案', () => download(filename(record.filename), file), 'secondary', !file))
}

export function recordEditor(record, editor) {
  const editors = { doc: documentEditor, sheet: sheetEditor, task: datedEditor, event: datedEditor, contact: contactEditor, asset: assetEditor }
  return el('div', { class: 'record-editor' },
    el('div', { class: 'editor-head' },
      el('div', { class: 'overline' }, types[record.kind]),
      inputField('名稱', record.title, value => editor.change({ title: value }), { class: 'name-input', maxlength: 250 }),
      inputField('專案', record.project, value => editor.change({ project: value }), {
        list: 'known-projects', maxlength: 120, placeholder: '輸入或選擇專案',
      }),
      el('datalist', { id: 'known-projects' }, editor.projects.map(project => el('option', { value: project }))),
      editor.indicator),
    editors[record.kind](record, editor))
}
