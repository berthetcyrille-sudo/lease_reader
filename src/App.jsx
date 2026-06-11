import { useState, useCallback, useRef } from 'react'
import { supabase } from './supabase.js'
import { SECTIONS, ALL_FIELDS, EXTRACTION_PROMPT } from './fields.js'
import { exportToExcel } from './export.js'
import './App.css'

// ─── helpers ────────────────────────────────────────────────────────────────

function toBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader()
    r.onload  = () => res(r.result.split(',')[1])
    r.onerror = () => rej(new Error('Lecture du fichier échouée'))
    r.readAsDataURL(file)
  })
}

function formatDate(iso) {
  if (!iso) return '—'
  return new Date(iso).toLocaleDateString('fr-FR', {
    day: '2-digit', month: 'short', year: 'numeric'
  })
}

// ─── sub-components ─────────────────────────────────────────────────────────

function DropZone({ onFile, disabled }) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef()

  const handle = useCallback(f => {
    if (!f) return
    const ext = f.name.split('.').pop().toLowerCase()
    if (!['pdf', 'docx'].includes(ext)) {
      alert('Format non supporté. Utilisez un PDF ou un DOCX.')
      return
    }
    onFile(f)
  }, [onFile])

  return (
    <div
      className={`drop-zone${dragging ? ' dragging' : ''}${disabled ? ' disabled' : ''}`}
      onClick={() => !disabled && inputRef.current?.click()}
      onDragOver={e => { e.preventDefault(); setDragging(true) }}
      onDragLeave={() => setDragging(false)}
      onDrop={e => { e.preventDefault(); setDragging(false); handle(e.dataTransfer.files[0]) }}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".pdf,.docx"
        style={{ display: 'none' }}
        onChange={e => handle(e.target.files[0])}
      />
      <div className="drop-icon">
        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
          <polyline points="14 2 14 8 20 8"/>
          <line x1="12" y1="18" x2="12" y2="12"/>
          <line x1="9" y1="15" x2="15" y2="15"/>
        </svg>
      </div>
      <div className="drop-title">Déposez un bail ici</div>
      <div className="drop-sub">PDF ou DOCX · cliquez pour parcourir</div>
    </div>
  )
}

function ProgressBar({ active }) {
  return (
    <div className="progress-track">
      <div className={`progress-bar${active ? ' active' : ''}`} />
    </div>
  )
}

function FieldCard({ label, value }) {
  const empty = !value
  return (
    <div className="field-card">
      <div className="field-label">{label}</div>
      <div className={`field-value${empty ? ' empty' : ''}`}>
        {value || 'Non renseigné'}
      </div>
    </div>
  )
}

function ResultsView({ data }) {
  return (
    <div className="results">
      {SECTIONS.map(sec => (
        <div key={sec.id} className="section-block">
          <div className="section-title">{sec.label}</div>
          <div className="fields-grid">
            {sec.fields.map(f => (
              <FieldCard key={f.key} label={f.label} value={data[f.key]} />
            ))}
          </div>
        </div>
      ))}
    </div>
  )
}

function HistoryPanel({ history, onSelect, activeId }) {
  if (!history.length) return (
    <div className="history-empty">Aucune extraction sauvegardée</div>
  )
  return (
    <div className="history-list">
      {history.map(item => (
        <button
          key={item.id}
          className={`history-item${item.id === activeId ? ' active' : ''}`}
          onClick={() => onSelect(item)}
        >
          <div className="history-name">{item.file_name}</div>
          <div className="history-meta">
            {item.data?.preneur || '—'} · {formatDate(item.created_at)}
          </div>
        </button>
      ))}
    </div>
  )
}

// ─── main component ──────────────────────────────────────────────────────────

export default function App() {
  const [file,        setFile]        = useState(null)
  const [loading,     setLoading]     = useState(false)
  const [status,      setStatus]      = useState('')
  const [error,       setError]       = useState('')
  const [data,        setData]        = useState(null)
  const [activeId,    setActiveId]    = useState(null)
  const [history,     setHistory]     = useState([])
  const [histLoaded,  setHistLoaded]  = useState(false)
  const [tab,         setTab]         = useState('extract') // 'extract' | 'history'

  // ── load history ──
  async function loadHistory() {
    if (histLoaded) return
    const { data: rows, error: err } = await supabase
      .from('extractions')
      .select('id, file_name, created_at, data')
      .order('created_at', { ascending: false })
      .limit(50)
    if (!err) { setHistory(rows || []); setHistLoaded(true) }
  }

  function switchTab(t) {
    setTab(t)
    if (t === 'history') loadHistory()
  }

  // ── extraction ──
  async function handleExtract() {
    if (!file || loading) return
    setLoading(true)
    setError('')
    setData(null)
    setStatus('Lecture du fichier…')

    try {
      const base64 = await toBase64(file)
      setStatus('Envoi à Claude…')

      const ext = file.name.split('.').pop().toLowerCase()
      const mediaType = ext === 'pdf'
        ? 'application/pdf'
        : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

      // Call Supabase Edge Function
      const { data: fnData, error: fnErr } = await supabase.functions.invoke('extract-lease', {
        body: { base64, mediaType, prompt: EXTRACTION_PROMPT }
      })

      if (fnErr) throw new Error(fnErr.message)
      if (fnData?.error) throw new Error(fnData.error)

      const extracted = fnData.result
      setData(extracted)
      setStatus('Extraction terminée.')

      // Save to Supabase
      setStatus('Sauvegarde en base…')
      const { data: saved, error: saveErr } = await supabase
        .from('extractions')
        .insert({ file_name: file.name, data: extracted })
        .select()
        .single()

      if (!saveErr && saved) {
        setActiveId(saved.id)
        setHistory(prev => [saved, ...prev])
      }

      setStatus('')
    } catch (e) {
      setError(e.message)
      setStatus('')
    } finally {
      setLoading(false)
    }
  }

  function handleFileSelect(f) {
    setFile(f)
    setData(null)
    setError('')
    setActiveId(null)
  }

  function handleHistorySelect(item) {
    setData(item.data)
    setActiveId(item.id)
    setFile(null)
    setTab('extract')
  }

  function handleClear() {
    setFile(null)
    setData(null)
    setError('')
    setActiveId(null)
  }

  return (
    <div className="app">
      {/* ── sidebar ── */}
      <aside className="sidebar">
        <div className="sidebar-logo">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
            <polyline points="14 2 14 8 20 8"/>
            <line x1="16" y1="13" x2="8" y2="13"/>
            <line x1="16" y1="17" x2="8" y2="17"/>
          </svg>
          <span>Lease Reader</span>
        </div>

        <nav className="sidebar-nav">
          <button
            className={`nav-item${tab === 'extract' ? ' active' : ''}`}
            onClick={() => switchTab('extract')}
          >
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
            </svg>
            Extraire
          </button>
          <button
            className={`nav-item${tab === 'history' ? ' active' : ''}`}
            onClick={() => switchTab('history')}
          >
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <polyline points="12 8 12 12 14 14"/>
              <path d="M3.05 11a9 9 0 1 0 .5-4"/>
              <polyline points="3 3 3 7 7 7"/>
            </svg>
            Historique
            {history.length > 0 && <span className="badge">{history.length}</span>}
          </button>
        </nav>

        {tab === 'history' && (
          <HistoryPanel
            history={history}
            onSelect={handleHistorySelect}
            activeId={activeId}
          />
        )}
      </aside>

      {/* ── main ── */}
      <main className="main">
        <header className="topbar">
          <h1 className="page-title">
            {data
              ? (data.immeuble || data.adresse || file?.name || 'Résultat')
              : 'Extraction de bail'}
          </h1>
          {data && (
            <div className="topbar-actions">
              <button className="btn" onClick={() => exportToExcel(data, file?.name || activeId || 'bail')}>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                  <polyline points="7 10 12 15 17 10"/>
                  <line x1="12" y1="15" x2="12" y2="3"/>
                </svg>
                Excel
              </button>
              <button className="btn danger-outline" onClick={handleClear}>
                Nouveau
              </button>
            </div>
          )}
        </header>

        <div className="content">
          {!data && (
            <>
              <DropZone onFile={handleFileSelect} disabled={loading} />

              {file && !loading && (
                <div className="file-row">
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                    <polyline points="14 2 14 8 20 8"/>
                  </svg>
                  <span className="file-name">{file.name}</span>
                  <span className="file-size">({(file.size / 1024).toFixed(0)} Ko)</span>
                  <button className="btn primary" onClick={handleExtract}>
                    Extraire les données
                  </button>
                  <button className="btn" onClick={handleClear} style={{ marginLeft: 4 }}>✕</button>
                </div>
              )}

              {loading && (
                <div className="loading-block">
                  <ProgressBar active={loading} />
                  <div className="status-msg">{status}</div>
                </div>
              )}

              {error && (
                <div className="error-msg">
                  <strong>Erreur :</strong> {error}
                </div>
              )}
            </>
          )}

          {data && <ResultsView data={data} />}
        </div>
      </main>
    </div>
  )
}
