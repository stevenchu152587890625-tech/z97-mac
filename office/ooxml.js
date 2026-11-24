import { columnName, evaluateSheet } from './grid.js'
import { markdown } from './formats.js'
import { maxColumns, maxRows } from './storage.js'

export const wordType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
export const excelType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const partNamespace = 'http://schemas.openxmlformats.org/package/2006/'
const relNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const sheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const encode = new TextEncoder()
const decode = new TextDecoder()
const maxPartSize = 24 * 1024 * 1024
const crcTable = Uint32Array.from({ length: 256 }, (_, value) => {
  for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1
  return value >>> 0
})

function crc32(bytes) {
  let checksum = 0xffffffff
  for (const byte of bytes) checksum = crcTable[(checksum ^ byte) & 255] ^ (checksum >>> 8)
  return (checksum ^ 0xffffffff) >>> 0
}

function header(size, longFields, shortFields) {
  const view = new DataView(new ArrayBuffer(size))
  for (const [position, value] of longFields) view.setUint32(position, value, true)
  for (const [position, value] of shortFields) view.setUint16(position, value, true)
  return view.buffer
}

function zip(parts, type) {
  const files = []
  const directory = []
  let offset = 0
  let directorySize = 0
  for (const [path, contents] of Object.entries(parts)) {
    const name = encode.encode(path)
    const data = encode.encode(xmlHeader + contents)
    const checksum = crc32(data)
    files.push(header(30, [[0, 0x04034b50], [14, checksum], [18, data.length], [22, data.length]],
      [[4, 20], [12, 33], [26, name.length]]), name, data)
    directory.push(header(46, [[0, 0x02014b50], [16, checksum], [20, data.length], [24, data.length], [42, offset]],
      [[4, 20], [6, 20], [14, 33], [28, name.length]]), name)
    directorySize += 46 + name.length
    offset += 30 + name.length + data.length
  }
  const count = directory.length / 2
  return new Blob([...files, ...directory, header(22, [[0, 0x06054b50], [12, directorySize], [16, offset]],
    [[8, count], [10, count]])], { type })
}

const xmlEscape = value => String(value).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

function relationships(type, target) {
  return '<Relationships xmlns="' + partNamespace + 'relationships"><Relationship Id="rId1" Type="' +
    relNamespace + '/' + type + '" Target="' + target + '"/></Relationships>'
}

function contentTypes(parts) {
  return '<Types xmlns="' + partNamespace + 'content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    Object.entries(parts).map(([path, type]) => '<Override PartName="/' + path + '" ContentType="' + type + '"/>').join('') +
    '</Types>'
}

function wordRun(text, properties = '') {
  return '<w:r>' + (properties ? '<w:rPr>' + properties + '</w:rPr>' : '') +
    '<w:t xml:space="preserve">' + xmlEscape(text) + '</w:t></w:r>'
}

function wordRuns(node, style = '') {
  if (node.nodeType === Node.TEXT_NODE) return wordRun(node.textContent, style)
  const tags = { STRONG: '<w:b/>', EM: '<w:i/>', CODE: '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>' }
  return Array.from(node.childNodes, child => wordRuns(child, style + (tags[node.tagName] ?? ''))).join('')
}

function wordParagraph(content, options = '') {
  return '<w:p>' + (options ? '<w:pPr>' + options + '</w:pPr>' : '') + content + '</w:p>'
}

export function writeWord(title, body) {
  const text = [wordParagraph(wordRun(title, '<w:b/><w:sz w:val="38"/>'))]
  for (const block of markdown(body).children) {
    if (/^H[1-6]$/.test(block.tagName)) {
      const level = +block.tagName[1]
      text.push(wordParagraph(wordRuns(block, '<w:b/><w:sz w:val="' + (36 - 2 * level) + '"/>'),
        '<w:outlineLvl w:val="' + (level - 1) + '"/>'))
    } else if (block.tagName === 'UL' || block.tagName === 'OL') {
      const start = block.start || 1
      Array.from(block.children).forEach((item, i) => {
        const marker = block.tagName === 'UL' ? '• ' : (start + i) + '. '
        text.push(wordParagraph(wordRun(marker) + wordRuns(item), '<w:ind w:left="340"/>'))
      })
    } else if (block.tagName === 'PRE') {
      for (const line of block.textContent.split('\n')) {
        text.push(wordParagraph(wordRun(line, '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>')))
      }
    } else if (block.tagName === 'HR') {
      text.push(wordParagraph(wordRun('────────────────')))
    } else {
      text.push(wordParagraph(wordRuns(block), block.tagName === 'BLOCKQUOTE' ? '<w:ind w:left="480"/>' : ''))
    }
  }
  const documentPath = 'word/document.xml'
  return zip({
    '[Content_Types].xml': contentTypes({ [documentPath]: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' }),
    '_rels/.rels': relationships('officeDocument', documentPath),
    [documentPath]: '<w:document xmlns:w="' + wordNamespace + '"><w:body>' + text.join('') + '</w:body></w:document>',
  }, wordType)
}

const supportedFunctions = new Set(['SUM', 'AVG', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'ROUND'])
function portableFormula(source) {
  if (!/^[\w \t.,():^+\-*/]+$/.test(source)) return ''
  for (const match of source.matchAll(/([A-Z][A-Z0-9]*)\s*\(/gi)) {
    if (!supportedFunctions.has(match[1].toUpperCase())) return ''
  }
  return source.replace(/\bAVG\s*\(/gi, 'AVERAGE(')
}

function excelNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (/^[+-]?(?:0|[1-9]\d{0,13})(?:\.\d+)?(?:e[+-]?\d+)?$/i.test(value) &&
      Number.isFinite(Number(value)) && value.replace(/[^\d]/g, '').length <= 15) return Number(value)
  return null
}

export function writeExcel(rows) {
  const evaluate = evaluateSheet(rows)
  const rowXML = rows.map((row, r) => {
    const cells = row.map((value, c) => {
      if (!value) return ''
      const address = ' r="' + columnName(c) + (r + 1) + '"'
      const formula = value.startsWith('=')
      const literal = value.startsWith("'")
      const evaluated = formula ? evaluate(r, c) : literal ? value.slice(1) : value
      const numeric = literal ? null : excelNumber(evaluated)
      const exported = formula && typeof evaluated === 'number' ? portableFormula(value.slice(1)) : ''
      if (exported) return '<c' + address + '><f>' + xmlEscape(exported) + '</f><v>' + numeric + '</v></c>'
      if (numeric !== null) return '<c' + address + '><v>' + numeric + '</v></c>'
      return '<c' + address + ' t="inlineStr"><is><t xml:space="preserve">' + xmlEscape(evaluated) + '</t></is></c>'
    }).join('')
    return cells ? '<row r="' + (r + 1) + '">' + cells + '</row>' : ''
  }).join('')
  const book = 'xl/workbook.xml'
  const sheet = 'xl/worksheets/sheet1.xml'
  return zip({
    '[Content_Types].xml': contentTypes({
      [book]: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
      [sheet]: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
    }),
    '_rels/.rels': relationships('officeDocument', book),
    'xl/_rels/workbook.xml.rels': relationships('worksheet', 'worksheets/sheet1.xml'),
    [book]: '<workbook xmlns="' + sheetNamespace + '" xmlns:r="' + relNamespace + '">' +
      '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    [sheet]: '<worksheet xmlns="' + sheetNamespace + '"><sheetData>' + rowXML + '</sheetData></worksheet>',
  }, excelType)
}

function badZip() {
  return new Error('Office 檔案的壓縮結構無效，或使用此工作台未支援的壓縮格式。')
}

async function unpack(file) {
  const source = await file.arrayBuffer()
  const view = new DataView(source)
  const u16 = position => view.getUint16(position, true)
  const u32 = position => view.getUint32(position, true)
  const slice = (from, length) => new Uint8Array(source, from, length)
  let end = source.byteLength - 22
  const minEnd = Math.max(0, end - 65535)
  while (end >= minEnd && (u32(end) !== 0x06054b50 || end + 22 + u16(end + 20) !== source.byteLength)) end--
  if (end < minEnd || u16(end + 4) || u16(end + 6) || u16(end + 8) !== u16(end + 10)) throw badZip()
  const count = u16(end + 10)
  if (count > 4096) throw badZip()
  const entries = new Map()
  let offset = u32(end + 16)
  if (offset + u32(end + 12) > end) throw badZip()
  for (let i = 0; i < count; i++) {
    if (offset + 46 > end || u32(offset) !== 0x02014b50) throw badZip()
    const nameLength = u16(offset + 28)
    const size = 46 + nameLength + u16(offset + 30) + u16(offset + 32)
    if (offset + size > end) throw badZip()
    const name = decode.decode(slice(offset + 46, nameLength))
    const entry = {
      flags: u16(offset + 8), method: u16(offset + 10), crc: u32(offset + 16), compressed: u32(offset + 20),
      original: u32(offset + 24), local: u32(offset + 42),
    }
    if (entries.has(name)) throw badZip()
    entries.set(name, entry)
    offset += size
  }
  return async path => {
    const item = entries.get(path)
    if (!item) return null
    if (item.original > maxPartSize || item.flags & 1 || ![0, 8].includes(item.method) ||
        item.local + 30 > end || u32(item.local) !== 0x04034b50) throw badZip()
    const start = item.local + 30 + u16(item.local + 26) + u16(item.local + 28)
    if (start + item.compressed > end) throw badZip()
    let bytes = slice(start, item.compressed)
    if (item.method === 8) {
      if (typeof DecompressionStream !== 'function') throw new Error('此瀏覽器不支援解壓縮 Office 檔案。')
      let stream
      try {
        stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader()
      } catch {
        throw new Error('此瀏覽器不支援解壓縮 Office 檔案。')
      }
      bytes = new Uint8Array(item.original)
      let read = 0
      try {
        while (true) {
          const chunk = await stream.read()
          if (chunk.done) break
          if (read + chunk.value.byteLength > bytes.length) throw badZip()
          bytes.set(chunk.value, read)
          read += chunk.value.byteLength
        }
      } finally {
        await stream.cancel().catch(() => {})
      }
      if (read !== item.original) throw badZip()
    }
    if (bytes.length !== item.original || crc32(bytes) !== item.crc) throw badZip()
    const decoder = bytes[0] === 255 && bytes[1] === 254 ? new TextDecoder('utf-16le')
      : bytes[0] === 254 && bytes[1] === 255 ? new TextDecoder('utf-16be') : decode
    const xml = decoder.decode(bytes)
    if (/<!DOCTYPE|<!ENTITY/i.test(xml)) throw badZip()
    const parsed = new DOMParser().parseFromString(xml, 'application/xml')
    if (parsed.getElementsByTagName('parsererror').length) throw badZip()
    return parsed
  }
}

const elements = (node, name) => Array.from(node?.getElementsByTagNameNS('*', name) ?? [])
const first = (node, name) => node?.getElementsByTagNameNS('*', name)?.[0]
function targetPath(base, relation) {
  const target = relation?.getAttribute('Target')
  if (!target || relation.getAttribute('TargetMode') === 'External') throw badZip()
  const url = new URL(target, 'https://archive.invalid/' + base)
  if (url.origin !== 'https://archive.invalid' || url.search || url.hash) throw badZip()
  return decodeURIComponent(url.pathname.slice(1))
}

async function mainPart(read) {
  const relations = await read('_rels/.rels')
  return targetPath('', elements(relations, 'Relationship').find(item => item.getAttribute('Type')?.endsWith('/officeDocument')))
}

export async function readWord(file) {
  const read = await unpack(file)
  const path = await mainPart(read)
  const document = await read(path)
  const body = first(document, 'body')
  if (!body) throw new Error('找不到 DOCX 的文件內容。')
  let characters = 0
  const lines = elements(body, 'p').map(paragraph => {
    const value = elements(paragraph, 'r').map(run => {
      const content = Array.from(run.children, node => node.localName === 't' ? node.textContent :
        node.localName === 'tab' ? '\t' : node.localName === 'br' ? '\n' : '').join('')
      if (!content.trim()) return content
      if (first(run, 'b')) return '**' + content + '**'
      if (first(run, 'i')) return '*' + content + '*'
      return content
    }).join('')
    const level = first(paragraph, 'outlineLvl')?.getAttributeNS(wordNamespace, 'val')
    const style = first(paragraph, 'pStyle')?.getAttributeNS(wordNamespace, 'val')
    const heading = /^heading([1-6])$/i.exec(style ?? '')?.[1]
    const prefix = value && (heading || (level !== undefined && level !== null && /^[0-5]$/.test(level)))
      ? '#'.repeat(heading ? Number(heading) : Number(level) + 1) + ' ' : ''
    characters += prefix.length + value.length + 2
    if (characters > 2_000_000) throw new Error('文件太大，可改從檔案頁匯入為附件。')
    return prefix + value
  })
  return lines.join('\n\n')
}

export async function readExcel(file) {
  const read = await unpack(file)
  const path = await mainPart(read)
  const book = await read(path)
  const sheets = elements(book, 'sheet')
  const sheet = sheets[0]
  if (!sheet) throw new Error('找不到 XLSX 工作表。')
  const dirEnd = path.lastIndexOf('/') + 1
  const links = await read(path.slice(0, dirEnd) + '_rels/' + path.slice(dirEnd) + '.rels')
  const relationships = elements(links, 'Relationship')
  const relationID = Array.from(sheet.attributes).find(attribute => attribute.localName === 'id')?.value
  const relation = relationships.find(item => item.getAttribute('Id') === relationID && item.getAttribute('Type')?.endsWith('/worksheet'))
  if (!relation) throw new Error('第一個分頁不是可匯入的工作表。')
  const document = await read(targetPath(path, relation))
  if (!first(document, 'sheetData')) throw new Error('找不到 XLSX 儲存格。')
  const stringsRelation = relationships.find(item => item.getAttribute('Type')?.endsWith('/sharedStrings'))
  const strings = stringsRelation ? elements(await read(targetPath(path, stringsRelation)), 'si')
    .map(item => elements(item, 't').map(run => run.textContent).join('')) : []
  const rows = []
  for (const cell of elements(document, 'c')) {
    const address = /^([A-Z]+)([1-9]\d*)$/.exec(cell.getAttribute('r') ?? '')
    const formula = first(cell, 'f')?.textContent ?? ''
    const scalar = first(cell, 'v')?.textContent ?? ''
    const inline = cell.getAttribute('t') === 'inlineStr'
    if (!address || (!formula && !scalar && !inline)) continue
    const col = [...address[1]].reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1
    const row = Number(address[2]) - 1
    if (row >= maxRows || col >= maxColumns) throw new Error('第一張表使用的範圍超過 ' + maxRows + ' 列 × ' + maxColumns + ' 欄。')
    let value = scalar
    const type = cell.getAttribute('t')
    if (type === 's') value = strings[Number(scalar)] ?? ''
    if (type === 'b') value = scalar === '1' ? 'TRUE' : 'FALSE'
    if (inline) value = elements(cell, 't').map(text => text.textContent).join('')
    if (formula && portableFormula(formula)) value = '=' + formula
    else if (/^[=']/.test(value)) value = "'" + value
    if (value.length > 20000) throw new Error('工作表含有超過 20,000 字元的儲存格。')
    while (rows.length <= row) rows.push([])
    while (rows[row].length < col) rows[row].push('')
    rows[row][col] = value
  }
  return { rows: rows.length ? rows : [['']], sheetCount: sheets.length }
}
