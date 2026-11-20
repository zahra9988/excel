import { useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx-js-style'

const THIN = { style: 'thin', color: { rgb: '000000' } }
const BORDER = { top: THIN, bottom: THIN, left: THIN, right: THIN }
const HEADER_STYLE = { font: { bold: true }, border: BORDER }
const CELL_STYLE = { border: BORDER }

const STORAGE_KEY = 'sheet-builder-v1'
const uid = () => Math.random().toString(36).slice(2, 10)

/* Column types the user can choose from */
const TYPES = [
  { value: 'text', label: 'Text', icon: 'Aa' },
  { value: 'number', label: 'Number', icon: '12' },
  { value: 'currency', label: 'Currency', icon: '$' },
  { value: 'percent', label: 'Percentage', icon: '%' },
  { value: 'date', label: 'Date', icon: '📅' },
  { value: 'time', label: 'Time', icon: '🕒' },
  { value: 'datetime', label: 'Date & Time', icon: '🗓' },
  { value: 'yesno', label: 'Yes / No', icon: '✓' },
  { value: 'dropdown', label: 'Dropdown list', icon: '▾' },
  { value: 'email', label: 'Email', icon: '@' },
  { value: 'phone', label: 'Phone', icon: '☎' },
]
const typeInfo = (t) => TYPES.find((x) => x.value === t) || TYPES[0]

const newColumn = (name, type = 'text') => ({ id: uid(), name, type, options: '' })

const newSheet = (name) => ({
  id: uid(),
  name,
  columns: [newColumn('Parameter 1'), newColumn('Parameter 2'), newColumn('Parameter 3')],
  rows: [],
})

function freshState() {
  const first = newSheet('Sheet1')
  return { fileName: 'My Sheet', sheets: [first], activeId: first.id, submitted: false }
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (raw) {
      const s = JSON.parse(raw)
      // Older single-sheet saves: wrap columns/rows into one sheet
      if (!s.sheets) {
        const sh = { id: uid(), name: 'Sheet1', columns: s.columns || [], rows: s.rows || [] }
        s.sheets = [sh]
        s.activeId = sh.id
        delete s.columns; delete s.rows
      }
      s.sheets = s.sheets.map((sh) => ({
        ...sh,
        columns: sh.columns.map((c) => ({ type: 'text', options: '', ...c })),
      }))
      if (!s.sheets.some((sh) => sh.id === s.activeId)) s.activeId = s.sheets[0].id
      return { fileName: 'My Sheet', submitted: false, ...s }
    }
  } catch {}
  return freshState()
}

// Excel sheet names: max 31 chars, none of : \ / ? * [ ]
const cleanSheetName = (n) => n.replace(/[:\\/?*[\]]/g, '').slice(0, 31)

const parseOptions = (str) =>
  (str || '').split(',').map((s) => s.trim()).filter(Boolean)

/* ---------- Excel conversion helpers ---------- */
const EPOCH = Date.UTC(1899, 11, 30)
function dateSerial(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s)
  if (!m) return null
  return (Date.UTC(+m[1], +m[2] - 1, +m[3]) - EPOCH) / 86400000
}
function timeSerial(s) {
  const m = /(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s)
  if (!m) return null
  return (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400
}
function toExcelCell(raw, type) {
  const s = (raw ?? '').toString().trim()
  if (s === '') return null
  const num = Number(s.replace(/,/g, ''))
  switch (type) {
    case 'number':
      return isFinite(num) ? { t: 'n', v: num } : { t: 's', v: s }
    case 'currency':
      return isFinite(num) ? { t: 'n', v: num, z: '#,##0.00' } : { t: 's', v: s }
    case 'percent':
      return isFinite(num) ? { t: 'n', v: num / 100, z: '0.00%' } : { t: 's', v: s }
    case 'date': {
      const d = dateSerial(s)
      return d != null ? { t: 'n', v: d, z: 'dd-mmm-yyyy' } : { t: 's', v: s }
    }
    case 'time': {
      const t = timeSerial(s)
      return t != null ? { t: 'n', v: t, z: 'hh:mm AM/PM' } : { t: 's', v: s }
    }
    case 'datetime': {
      const d = dateSerial(s), t = timeSerial(s)
      return d != null ? { t: 'n', v: d + (t || 0), z: 'dd-mmm-yyyy hh:mm AM/PM' } : { t: 's', v: s }
    }
    default:
      return { t: 's', v: s }
  }
}

/* ---------- Display formatting ---------- */
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
function fmtDate(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s)
  return m ? `${m[3]}-${MONTHS[+m[2] - 1]}-${m[1]}` : s
}
function fmtTime(s) {
  const m = /(\d{1,2}):(\d{2})$/.exec(s)
  if (!m) return s
  const h = +m[1]
  return `${((h + 11) % 12) + 1}:${m[2]} ${h < 12 ? 'AM' : 'PM'}`
}
function displayValue(raw, type) {
  const s = (raw ?? '').toString()
  if (!s) return ''
  switch (type) {
    case 'date': return fmtDate(s)
    case 'time': return fmtTime(s)
    case 'datetime': {
      const [d, t] = s.split('T')
      return t ? `${fmtDate(d)} ${fmtTime(t)}` : fmtDate(d)
    }
    case 'currency': {
      const n = Number(s.replace(/,/g, ''))
      return isFinite(n) ? n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : s
    }
    case 'percent': return `${s}%`
    default: return s
  }
}

/* ---------- Validation ---------- */
function validate(raw, type) {
  const s = (raw ?? '').toString().trim()
  if (!s) return ''
  if (['number', 'currency', 'percent'].includes(type) && !isFinite(Number(s.replace(/,/g, ''))))
    return 'Must be a number'
  if (type === 'email' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s)) return 'Invalid email'
  if (type === 'phone' && !/^[+\d][\d\s\-()]{4,}$/.test(s)) return 'Invalid phone number'
  return ''
}

export default function App() {
  const [state, setState] = useState(loadState)
  const { fileName, sheets, activeId, submitted } = state
  const sheet = sheets.find((s) => s.id === activeId) || sheets[0]
  const { columns, rows } = sheet
  const [newCol, setNewCol] = useState('')
  const [newType, setNewType] = useState('text')
  const [draft, setDraft] = useState({})
  const [touched, setTouched] = useState(false)
  const [editingRow, setEditingRow] = useState(null)
  const [toast, setToast] = useState('')
  const firstInputRef = useRef(null)

  useEffect(() => {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)) } catch {}
  }, [state])

  useEffect(() => {
    if (!toast) return
    const t = setTimeout(() => setToast(''), 2400)
    return () => clearTimeout(t)
  }, [toast])

  const update = (patch) => setState((s) => ({ ...s, ...patch }))
  const updateSheet = (patch) =>
    setState((s) => ({ ...s, sheets: s.sheets.map((sh) => (sh.id === sheet.id ? { ...sh, ...patch } : sh)) }))

  /* ---------- Sheets ---------- */
  const resetEntry = () => { setDraft({}); setEditingRow(null); setTouched(false) }
  const selectSheet = (id) => { if (id !== sheet.id) { resetEntry(); update({ activeId: id }) } }
  const addSheet = () => {
    let n = sheets.length + 1
    while (sheets.some((s) => s.name.toLowerCase() === `sheet${n}`)) n++
    const sh = newSheet(`Sheet${n}`)
    resetEntry()
    update({ sheets: [...sheets, sh], activeId: sh.id })
  }
  const renameSheet = (id, name) => {
    const clean = cleanSheetName(name).trim()
    if (!clean) return setToast('Sheet name cannot be empty')
    if (sheets.some((s) => s.id !== id && s.name.toLowerCase() === clean.toLowerCase()))
      return setToast('A sheet with that name already exists')
    update({ sheets: sheets.map((s) => (s.id === id ? { ...s, name: clean } : s)) })
  }
  const deleteSheet = (id) => {
    if (sheets.length === 1) return setToast('A workbook needs at least one sheet')
    const sh = sheets.find((s) => s.id === id)
    if (sh.rows.length && !confirm(`Delete "${sh.name}" and its ${sh.rows.length} row(s)?`)) return
    const remaining = sheets.filter((s) => s.id !== id)
    if (id === sheet.id) resetEntry()
    update({ sheets: remaining, activeId: id === sheet.id ? remaining[0].id : sheet.id })
  }

  /* ---------- Columns ---------- */
  const addColumn = (e) => {
    e?.preventDefault()
    const name = newCol.trim() || `Parameter ${columns.length + 1}`
    updateSheet({ columns: [...columns, newColumn(name, newType)] })
    setNewCol('')
    setNewType('text')
  }
  const patchColumn = (id, patch) =>
    updateSheet({ columns: columns.map((c) => (c.id === id ? { ...c, ...patch } : c)) })
  const changeType = (col, type) => {
    const hasValues = rows.some((r) => (r[col.id] ?? '').toString().trim())
    if (hasValues && !confirm(`Changing "${col.name}" to ${typeInfo(type).label} will clear its existing values. Continue?`)) return
    updateSheet({
      columns: columns.map((c) => (c.id === col.id ? { ...c, type } : c)),
      rows: hasValues ? rows.map((r) => ({ ...r, [col.id]: '' })) : rows,
    })
    setDraft((d) => ({ ...d, [col.id]: '' }))
  }
  const removeColumn = (id) => {
    if (columns.length === 1) return setToast('At least one column is required')
    updateSheet({
      columns: columns.filter((c) => c.id !== id),
      rows: rows.map(({ [id]: _, ...rest }) => rest),
    })
  }
  const moveColumn = (idx, dir) => {
    const j = idx + dir
    if (j < 0 || j >= columns.length) return
    const next = [...columns]
    ;[next[idx], next[j]] = [next[j], next[idx]]
    updateSheet({ columns: next })
  }

  /* ---------- Rows ---------- */
  const draftEmpty = columns.every((c) => !(draft[c.id] ?? '').toString().trim())
  const errors = Object.fromEntries(columns.map((c) => [c.id, validate(draft[c.id], c.type)]))
  const hasErrors = Object.values(errors).some(Boolean)

  const saveRow = (e) => {
    e?.preventDefault()
    setTouched(true)
    if (draftEmpty) return setToast('Enter at least one value')
    if (hasErrors) return setToast('Please fix the highlighted fields')
    const values = Object.fromEntries(columns.map((c) => [c.id, (draft[c.id] ?? '').toString().trim()]))
    if (editingRow) {
      updateSheet({ rows: rows.map((r) => (r.id === editingRow ? { ...r, ...values } : r)) })
      setToast('Row updated')
    } else {
      updateSheet({ rows: [...rows, { id: uid(), ...values }] })
      setToast(`Row ${rows.length + 1} added`)
    }
    setDraft({})
    setTouched(false)
    setEditingRow(null)
    firstInputRef.current?.focus()
  }
  const editRow = (row) => {
    setEditingRow(row.id)
    setDraft(Object.fromEntries(columns.map((c) => [c.id, row[c.id] ?? ''])))
    firstInputRef.current?.focus()
  }
  const cancelEdit = () => { setEditingRow(null); setDraft({}); setTouched(false) }
  const deleteRow = (id) => {
    updateSheet({ rows: rows.filter((r) => r.id !== id) })
    if (editingRow === id) cancelEdit()
  }

  /* ---------- Submit / Export ---------- */
  const submit = () => {
    if (!totalRows) return setToast('Add at least one row before submitting')
    update({ submitted: true })
    window.scrollTo({ top: 0, behavior: 'smooth' })
  }

  const buildSheet = ({ columns, rows }) => {
    const header = columns.map((c) => c.name || 'Untitled')
    const ws = {}
    header.forEach((h, ci) => {
      ws[XLSX.utils.encode_cell({ r: 0, c: ci })] = { t: 's', v: h, s: HEADER_STYLE }
    })
    rows.forEach((r, ri) => {
      columns.forEach((c, ci) => {
        // Empty cells still get a cell so the border grid is complete
        const cell = toExcelCell(r[c.id], c.type) || { t: 's', v: '' }
        const s = cell.z ? { ...CELL_STYLE, numFmt: cell.z } : { ...CELL_STYLE }
        ws[XLSX.utils.encode_cell({ r: ri + 1, c: ci })] = { ...cell, s }
      })
    })
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: columns.length - 1 } })
    ws['!cols'] = columns.map((c, i) => ({
      wch: Math.min(50, Math.max(12, header[i].length, ...rows.map((r) => displayValue(r[c.id], c.type).length)) + 2),
    }))
    return ws
  }

  const exportExcel = () => {
    const wb = XLSX.utils.book_new()
    const used = new Set()
    sheets.forEach((sh, i) => {
      let name = cleanSheetName(sh.name).trim() || `Sheet${i + 1}`
      let base = name, k = 2
      while (used.has(name.toLowerCase())) name = `${base.slice(0, 28)} (${k++})`
      used.add(name.toLowerCase())
      XLSX.utils.book_append_sheet(wb, buildSheet(sh), name)
    })
    const safe = (fileName || 'sheet').replace(/[\\/:*?"<>|]+/g, '').trim() || 'sheet'
    XLSX.writeFile(wb, `${safe}.xlsx`)
    setToast('Excel file downloaded')
  }

  const startOver = () => {
    if (!confirm('Clear all sheets, columns and values and start a new workbook?')) return
    setState(freshState())
    resetEntry()
  }

  const totalRows = sheets.reduce((n, s) => n + s.rows.length, 0)
  const filledCells = useMemo(
    () => rows.reduce((n, r) => n + columns.filter((c) => (r[c.id] ?? '').toString().trim()).length, 0),
    [rows, columns]
  )

  const tabs = (
    <SheetTabs
      sheets={sheets}
      activeId={sheet.id}
      onSelect={selectSheet}
      onAdd={submitted ? null : addSheet}
      onRename={submitted ? null : renameSheet}
      onDelete={submitted ? null : deleteSheet}
    />
  )

  return (
    <div className="page">
      <header className="topbar">
        <div className="brand">
          <span className="logo" aria-hidden>
            <svg viewBox="0 0 24 24" width="18" height="18"><path fill="currentColor" d="M4 3h16a1 1 0 0 1 1 1v16a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Zm1 2v4h6V5H5Zm8 0v4h6V5h-6Zm-8 6v4h6v-4H5Zm8 0v4h6v-4h-6Zm-8 6v2h6v-2H5Zm8 0v2h6v-2h-6Z"/></svg>
          </span>
          <span>Sheet Builder</span>
        </div>
        <div className="steps">
          <Step n={1} label="Columns" active={!submitted} done={submitted} />
          <span className="step-line" />
          <Step n={2} label="Values" active={!submitted} done={submitted} />
          <span className="step-line" />
          <Step n={3} label="Export" active={submitted} />
        </div>
      </header>

      <main className="container">
        <section className="card file-card">
          <label className="label" htmlFor="fname">File name</label>
          <div className="file-row">
            <input
              id="fname"
              className="input"
              value={fileName}
              onChange={(e) => update({ fileName: e.target.value })}
              placeholder="My Sheet"
            />
            <span className="ext">.xlsx</span>
          </div>
        </section>

        {tabs}

        {submitted ? (
          <section className="card">
            <div className="card-head">
              <div>
                <h2>Ready to export</h2>
                <p className="muted">
                  {sheets.length} sheet{sheets.length !== 1 && 's'} · {totalRows} row{totalRows !== 1 && 's'} in total
                </p>
                <p className="muted">
                  <strong>{sheet.name}:</strong> {rows.length} row{rows.length !== 1 && 's'} · {columns.length} column{columns.length !== 1 && 's'} · {filledCells} filled cells
                </p>
              </div>
              <div className="actions">
                <button className="btn ghost" onClick={() => update({ submitted: false })}>Back to edit</button>
                <button className="btn primary" onClick={exportExcel}>
                  <DownloadIcon /> Export Excel
                </button>
              </div>
            </div>
            {rows.length ? (
              <DataTable columns={columns} rows={rows} />
            ) : (
              <div className="empty">This sheet has no rows. It will be exported with its header row only.</div>
            )}
          </section>
        ) : (
          <>
            {/* Columns */}
            <section className="card">
              <div className="card-head">
                <div>
                  <h2><span className="num">1</span>Columns</h2>
                  <p className="muted">Name each column and choose what kind of value it holds.</p>
                </div>
              </div>

              <ul className="col-list">
                {columns.map((c, i) => (
                  <li key={c.id} className="col-block">
                    <div className="col-item">
                      <span className="col-letter">{XLSX.utils.encode_col(i)}</span>
                      <input
                        className="input"
                        value={c.name}
                        onChange={(e) => patchColumn(c.id, { name: e.target.value })}
                        placeholder="Column name"
                        aria-label={`Column ${i + 1} name`}
                      />
                      <TypeSelect value={c.type} onChange={(t) => changeType(c, t)} label={`Column ${i + 1} type`} />
                      <div className="icon-group">
                        <button className="icon-btn" onClick={() => moveColumn(i, -1)} disabled={i === 0} title="Move left">‹</button>
                        <button className="icon-btn" onClick={() => moveColumn(i, 1)} disabled={i === columns.length - 1} title="Move right">›</button>
                        <button className="icon-btn danger" onClick={() => removeColumn(c.id)} title="Remove column">×</button>
                      </div>
                    </div>
                    {c.type === 'dropdown' && (
                      <div className="col-options">
                        <input
                          className="input sm"
                          value={c.options}
                          onChange={(e) => patchColumn(c.id, { options: e.target.value })}
                          placeholder="Options, separated by commas — e.g. Pending, Approved, Rejected"
                          aria-label={`Column ${i + 1} dropdown options`}
                        />
                        <div className="chips">
                          {parseOptions(c.options).map((o) => <span key={o} className="chip">{o}</span>)}
                        </div>
                      </div>
                    )}
                  </li>
                ))}
              </ul>

              <form className="add-col" onSubmit={addColumn}>
                <input
                  className="input"
                  value={newCol}
                  onChange={(e) => setNewCol(e.target.value)}
                  placeholder="New column name"
                />
                <TypeSelect value={newType} onChange={setNewType} label="New column type" />
                <button className="btn secondary" type="submit">+ Add column</button>
              </form>
            </section>

            {/* Values */}
            <section className="card">
              <div className="card-head">
                <div>
                  <h2><span className="num">2</span>Values</h2>
                  <p className="muted">Fill in a value for each column, then add the row. Press Enter to add quickly.</p>
                </div>
                {editingRow && <span className="badge">Editing row {rows.findIndex((r) => r.id === editingRow) + 1}</span>}
              </div>

              <form className="entry" onSubmit={saveRow} noValidate>
                <div className="entry-grid">
                  {columns.map((c, i) => {
                    const err = touched || (draft[c.id] ?? '') !== '' ? errors[c.id] : ''
                    return (
                      <div key={c.id} className="field">
                        <label className="label" htmlFor={`f-${c.id}`}>
                          {c.name || 'Untitled'} <span className="type-hint">{typeInfo(c.type).label}</span>
                        </label>
                        <ValueInput
                          id={`f-${c.id}`}
                          inputRef={i === 0 ? firstInputRef : undefined}
                          column={c}
                          value={draft[c.id] ?? ''}
                          invalid={!!err}
                          onChange={(v) => setDraft((d) => ({ ...d, [c.id]: v }))}
                        />
                        {err && <span className="field-error">{err}</span>}
                      </div>
                    )
                  })}
                </div>
                <div className="entry-actions">
                  {editingRow && <button type="button" className="btn ghost" onClick={cancelEdit}>Cancel</button>}
                  <button type="submit" className="btn secondary" disabled={draftEmpty}>
                    {editingRow ? 'Save changes' : '+ Add row'}
                  </button>
                </div>
              </form>

              {rows.length > 0 ? (
                <DataTable columns={columns} rows={rows} onEdit={editRow} onDelete={deleteRow} editingRow={editingRow} />
              ) : (
                <div className="empty">No rows yet. Your entries will appear here.</div>
              )}
            </section>

            <div className="footer-bar">
              <button className="btn link" onClick={startOver}>Start over</button>
              <div className="footer-right">
                <span className="muted">
                  {rows.length} row{rows.length !== 1 && 's'} in {sheet.name}
                  {sheets.length > 1 && ` · ${totalRows} total`}
                </span>
                <button className="btn primary" onClick={submit} disabled={!totalRows}>Submit</button>
              </div>
            </div>
          </>
        )}
      </main>

      {toast && <div className="toast">{toast}</div>}
    </div>
  )
}

function SheetTabs({ sheets, activeId, onSelect, onAdd, onRename, onDelete }) {
  const [editing, setEditing] = useState(null)
  const [name, setName] = useState('')
  const startRename = (s) => { if (!onRename) return; setEditing(s.id); setName(s.name) }
  const commit = () => { if (editing) onRename(editing, name); setEditing(null) }

  return (
    <div className="sheet-bar">
      <div className="sheet-tabs" role="tablist">
        {sheets.map((s) => {
          const active = s.id === activeId
          return (
            <div
              key={s.id}
              role="tab"
              aria-selected={active}
              tabIndex={0}
              className={`sheet-tab${active ? ' active' : ''}`}
              onClick={() => onSelect(s.id)}
              onDoubleClick={() => startRename(s)}
              onKeyDown={(e) => e.key === 'Enter' && editing !== s.id && onSelect(s.id)}
              title={onRename ? 'Double-click to rename' : s.name}
            >
              {editing === s.id ? (
                <input
                  className="sheet-input"
                  value={name}
                  autoFocus
                  maxLength={31}
                  onChange={(e) => setName(e.target.value)}
                  onBlur={commit}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') commit()
                    if (e.key === 'Escape') setEditing(null)
                  }}
                  onClick={(e) => e.stopPropagation()}
                  aria-label="Sheet name"
                />
              ) : (
                <>
                  <span className="sheet-name">{s.name}</span>
                  {s.rows.length > 0 && <span className="sheet-count">{s.rows.length}</span>}
                </>
              )}
              {active && onRename && editing !== s.id && (
                <button className="sheet-icon" title="Rename sheet" aria-label={`Rename ${s.name}`}
                  onClick={(e) => { e.stopPropagation(); startRename(s) }}>✎</button>
              )}
              {onDelete && sheets.length > 1 && editing !== s.id && (
                <button className="sheet-icon danger" title="Delete sheet" aria-label={`Delete ${s.name}`}
                  onClick={(e) => { e.stopPropagation(); onDelete(s.id) }}>×</button>
              )}
            </div>
          )
        })}
      </div>
      {onAdd && (
        <button className="btn secondary sm" onClick={onAdd}>+ Add sheet</button>
      )}
    </div>
  )
}

function TypeSelect({ value, onChange, label }) {
  return (
    <div className="type-select">
      <span className="type-icon" aria-hidden>{typeInfo(value).icon}</span>
      <select className="input select" value={value} onChange={(e) => onChange(e.target.value)} aria-label={label}>
        {TYPES.map((t) => <option key={t.value} value={t.value}>{t.label}</option>)}
      </select>
    </div>
  )
}

function ValueInput({ id, inputRef, column, value, invalid, onChange }) {
  const cls = `input${invalid ? ' invalid' : ''}`
  const common = { id, ref: inputRef, className: cls, value, onChange: (e) => onChange(e.target.value) }
  switch (column.type) {
    case 'number':
      return <input {...common} inputMode="decimal" placeholder="0" />
    case 'currency':
      return (
        <div className="affix">
          <input {...common} inputMode="decimal" placeholder="0.00" />
        </div>
      )
    case 'percent':
      return (
        <div className="affix suffix" data-affix="%">
          <input {...common} inputMode="decimal" placeholder="0" />
        </div>
      )
    case 'date':
      return <input {...common} type="date" />
    case 'time':
      return <input {...common} type="time" />
    case 'datetime':
      return <input {...common} type="datetime-local" />
    case 'email':
      return <input {...common} type="email" placeholder="name@example.com" />
    case 'phone':
      return <input {...common} type="tel" placeholder="+92 300 1234567" />
    case 'yesno':
      return (
        <select {...common} className={`${cls} select`}>
          <option value="">Select…</option>
          <option value="Yes">Yes</option>
          <option value="No">No</option>
        </select>
      )
    case 'dropdown': {
      const opts = parseOptions(column.options)
      return (
        <select {...common} className={`${cls} select`} disabled={!opts.length}>
          <option value="">{opts.length ? 'Select…' : 'Add options in Columns'}</option>
          {opts.map((o) => <option key={o} value={o}>{o}</option>)}
        </select>
      )
    }
    default:
      return <input {...common} placeholder={`Enter ${column.name || 'value'}`} />
  }
}

function Step({ n, label, active, done }) {
  return (
    <div className={`step ${active ? 'active' : ''} ${done ? 'done' : ''}`}>
      <span className="step-dot">{done ? '✓' : n}</span>
      <span className="step-label">{label}</span>
    </div>
  )
}

function DataTable({ columns, rows, onEdit, onDelete, editingRow }) {
  return (
    <div className="table-wrap">
      <table className="table">
        <thead>
          <tr>
            <th className="rownum">#</th>
            {columns.map((c) => (
              <th key={c.id} className={['number', 'currency', 'percent'].includes(c.type) ? 'num-col' : ''}>
                {c.name || 'Untitled'}
                <span className="th-type">{typeInfo(c.type).label}</span>
              </th>
            ))}
            {onEdit && <th className="act" />}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={r.id} className={editingRow === r.id ? 'editing' : ''}>
              <td className="rownum">{i + 1}</td>
              {columns.map((c) => {
                const v = displayValue(r[c.id], c.type)
                return (
                  <td key={c.id} data-label={c.name || 'Untitled'} className={['number', 'currency', 'percent'].includes(c.type) ? 'num-col' : ''}>
                    {v ? (c.type === 'yesno' ? <span className={`pill ${v === 'Yes' ? 'yes' : 'no'}`}>{v}</span> : v) : <span className="dash">—</span>}
                  </td>
                )
              })}
              {onEdit && (
                <td className="act">
                  <button className="icon-btn" onClick={() => onEdit(r)} title="Edit row">✎</button>
                  <button className="icon-btn danger" onClick={() => onDelete(r.id)} title="Delete row">×</button>
                </td>
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function DownloadIcon() {
  return (
    <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden>
      <path fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" d="M12 4v11m0 0-4-4m4 4 4-4M5 19h14" />
    </svg>
  )
}
