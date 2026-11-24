import { maxRows, maxColumns } from './storage.js'

export function columnName(index) {
  let label = ''
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    label = String.fromCharCode(65 + (n - 1) % 26) + label
  }
  return label
}

export function parseDelimited(text, separator = ',') {
  text = text.replace(/^\uFEFF/, '')
  const rows = []
  let row = []
  let value = ''
  let quoted = false
  let closed = false
  const cell = () => {
    row.push(value)
    if (row.length > maxColumns) throw new Error('試算表最多支援 ' + maxColumns + ' 欄。')
    value = ''
    closed = false
  }
  const line = () => {
    cell()
    rows.push(row)
    if (rows.length > maxRows) throw new Error('試算表最多支援 ' + maxRows + ' 列。')
    row = []
  }
  for (let i = 0; i < text.length; i++) {
    const char = text[i]
    if (quoted) {
      if (char === '"' && text[i + 1] === '"') {
        value += '"'
        i++
      } else if (char === '"') {
        quoted = false
        closed = true
      } else {
        value += char
      }
    } else if (char === separator) {
      cell()
    } else if (char === '\r' || char === '\n') {
      if (char === '\r' && text[i + 1] === '\n') i++
      line()
    } else if (closed) {
      throw new Error('CSV 引號後應接分隔符號或換行。')
    } else if (char === '"' && !value) {
      quoted = true
    } else {
      value += char
    }
    if (value.length > 20000) throw new Error('儲存格超過 20,000 字元。')
  }
  if (quoted) throw new Error('CSV 裡有未關閉的引號。')
  if (row.length || value || closed || !rows.length) line()
  return rows
}

class FormulaError extends Error {}
const fail = code => { throw new FormulaError(code) }
const numberText = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i
const numeric = value => {
  if (value === '') return 0
  if (typeof value === 'number') return value
  if (numberText.test(value)) return Number(value)
  return fail('#VALUE!')
}

function coordinates(label) {
  const [, col, row] = /^([A-Z]+)([1-9]\d*)$/.exec(label) ?? []
  if (!col) return fail('#NAME?')
  const index = [...col].reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1
  const line = Number(row) - 1
  if (index >= maxColumns || line >= maxRows) return fail('#REF!')
  return [line, index]
}

function calculate(source, read) {
  if (source.length > 4096) return fail('#LIMIT!')
  const tokens = source.toUpperCase().match(/\d+(?:\.\d*)?(?:E[+-]?\d+)?|\.\d+(?:E[+-]?\d+)?|[A-Z]+[A-Z0-9]*|[^\s]/g) ?? []
  let index = 0
  let depth = 0
  const peek = () => tokens[index]
  const accept = token => {
    if (peek() !== token) return false
    index++
    return true
  }
  const expect = token => { if (!accept(token)) fail('#ERROR!') }
  const nested = callback => {
    if (++depth > 80) return fail('#LIMIT!')
    const value = callback()
    depth--
    return value
  }
  const sum = () => {
    let value = product()
    while (peek() === '+' || peek() === '-') {
      const operator = tokens[index++]
      const right = numeric(product())
      value = numeric(value) + (operator === '+' ? right : -right)
    }
    return value
  }
  const product = () => {
    let value = unary()
    while (peek() === '*' || peek() === '/') {
      const operator = tokens[index++]
      const right = numeric(unary())
      if (operator === '/' && right === 0) return fail('#DIV/0!')
      value = operator === '*' ? numeric(value) * right : numeric(value) / right
    }
    return value
  }
  const unary = () => {
    if (accept('+')) return nested(() => numeric(unary()))
    if (accept('-')) return nested(() => -numeric(unary()))
    const value = atom()
    return accept('^') ? numeric(value) ** nested(() => numeric(unary())) : value
  }
  const atom = () => {
    if (accept('(')) {
      const value = nested(sum)
      expect(')')
      return value
    }
    const token = tokens[index++]
    if (!token) return fail('#ERROR!')
    if (numberText.test(token)) return Number(token)
    if (accept('(')) {
      const values = []
      if (!accept(')')) {
        do {
          if (tokens[index + 1] === ':') {
            const start = coordinates(tokens[index++])
            index++
            const end = coordinates(tokens[index++] ?? '')
            for (let row = Math.min(start[0], end[0]); row <= Math.max(start[0], end[0]); row++) {
              for (let col = Math.min(start[1], end[1]); col <= Math.max(start[1], end[1]); col++) {
                values.push(read(row, col))
              }
            }
          } else {
            values.push(nested(sum))
          }
        } while (accept(','))
        expect(')')
      }
      const numbers = values.filter(value => typeof value === 'number' || numberText.test(value)).map(Number)
      const total = () => numbers.reduce((a, b) => a + b, 0)
      switch (token) {
        case 'SUM': return total()
        case 'AVG':
        case 'AVERAGE': return numbers.length ? total() / numbers.length : fail('#DIV/0!')
        case 'MIN': return numbers.reduce((a, b) => Math.min(a, b), numbers[0] ?? 0)
        case 'MAX': return numbers.reduce((a, b) => Math.max(a, b), numbers[0] ?? 0)
        case 'COUNT': return numbers.length
        case 'COUNTA': return values.filter(value => value !== '').length
        case 'ROUND': {
          const digits = numeric(values[1] ?? 0)
          if (!values.length || !Number.isInteger(digits) || Math.abs(digits) > 15) return fail('#VALUE!')
          const scale = 10 ** digits
          const value = numeric(values[0])
          return Math.sign(value) * Math.round(Math.abs(value) * scale) / scale
        }
        default: return fail('#NAME?')
      }
    }
    return read(...coordinates(token))
  }
  const value = sum()
  if (index !== tokens.length) return fail('#ERROR!')
  return value
}

export function evaluateSheet(rows) {
  const cache = new Map()
  const active = new Set()
  let operations = 0
  let budget = 2_000_000
  const read = (row, col) => {
    if (++operations > 200000 || --budget < 0) return fail('#LIMIT!')
    const key = row * maxColumns + col
    const previous = cache.get(key)
    if (previous instanceof FormulaError) throw previous
    if (cache.has(key)) return previous
    const raw = rows[row]?.[col] ?? ''
    if (raw.startsWith("'")) return raw.slice(1)
    if (!raw.startsWith('=')) return raw
    if (active.has(key)) return fail('#CYCLE!')
    if (active.size >= 200) return fail('#LIMIT!')
    active.add(key)
    try {
      const result = calculate(raw.slice(1), read)
      if (typeof result === 'number' && !Number.isFinite(result)) return fail('#NUM!')
      cache.set(key, result)
      return result
    } catch (error) {
      if (!(error instanceof FormulaError)) throw error
      cache.set(key, error)
      throw error
    } finally {
      active.delete(key)
    }
  }
  return (row, col) => {
    const raw = rows[row]?.[col] ?? ''
    if (raw.startsWith("'")) return raw.slice(1)
    if (!raw.startsWith('=')) return raw
    operations = 0
    try {
      return read(row, col)
    } catch (error) {
      if (error instanceof FormulaError) return error.message
      throw error
    }
  }
}

export function writeCSV(rows) {
  const evaluated = evaluateSheet(rows)
  const safe = value => {
    let text = String(value)
    // Protect downstream CSV readers from interpreting untrusted text as a formula.
    if (/^[\s\uFEFF]*[=+@-]|^[\t\r]/.test(text) && !numberText.test(text)) text = "'" + text
    return '"' + text.replace(/"/g, '""') + '"'
  }
  return '\uFEFF' + rows.map((row, r) => row.map((_, c) => safe(evaluated(r, c))).join(',')).join('\r\n')
}
