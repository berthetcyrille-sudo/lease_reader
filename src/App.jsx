import { useState, useCallback, useRef } from 'react'
import { supabase } from './supabase.js'
import { SECTIONS, ALL_FIELDS, EXTRACTION_PROMPT, AVENANT_PROMPT } from './fields.js'
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

function getMediaType(file) {
  return file.name.toLowerCase().endsWith('.pdf')
    ? 'application/pdf'
    : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
}

function formatDate(iso) {
  if (!iso) return '—'
  return new Date(iso).toLocaleDateString('fr-FR', { day: '2-digit', month: 'short', year: 'numeric' })
}

function similarity(a, b) {
  if (!a || !b) return 0
  const na = a.toLowerCase().trim()
  const nb = b.toLowerCase().trim()
  if (na === nb) return 1
  if (na.includes(nb) || nb.includes(na)) return 0.8
  const wordsA = na.split(/\s+/)
  const wordsB = nb.split(/\s+/)
  const common = wordsA.filter(w => wordsB.includes(w)).length
  return common / Math.max(wordsA.length, wordsB.length)
}

function findBestMatch(ref, bails) {
  if (!ref || !bails.length) return null
  let best = null, bestScore = 0
  for (const b of bails) {
    const d = b.data || {}
    const score =
      similarity(ref.preneur,  d.preneur)  * 0.4 +
      similarity(ref.bailleur, d.bailleur) * 0.2 +
      similarity(ref.adresse,  d.adresse)  * 0.2 +
      similarity(ref.immeuble, d.immeuble) * 0.2
    if (score > bestScore) { bestScore = score; best = b }
  }
  return bestScore > 0.3 ? { item: best, score: bestScore } : null
}

// ─── sub-components ─────────────────────────────────────────────────────────

function DropZone({ onFiles, disabled, multi }) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef()

  const handle = useCallback(files => {
    const valid = Array.from(files).filter(f => {
      const ext = f.name.split('.').pop().toLowerCase()
      return ['pdf', 'docx'].includes(ext)
    })
    if (valid.length) onFiles(valid)
    else alert('Format non supporté. Utilisez des fichiers PDF ou DOCX.')
  }, [onFiles])

  return (
    <div
      className={`drop-zone${dragging ? ' dragging' : ''}${disabled ? ' disabled' : ''}`}
      onClick={() => !disabled && inputRef.current?.click()}
      onDragOver={e => { e.preventDefault(); setDragging(true) }}
      onDragLeave={() => setDragging(false)}
      onDrop={e => { e.preventDefault(); setDragging(false); handle(e.dataTransfer.files) }}
    >
      <input ref={inputRef} type="file" accept=".pdf,.docx" multiple={multi} style={{ display: 'none' }}
        onChange={e => handle(e.target.files)} />
      <div className="drop-icon">
        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
          <polyline points="14 2 14 8 20 8"/>
          <line x1="12" y1="18" x2="12" y2="12"/>
          <line x1="9" y1="15" x2="15" y2="15"/>
        </svg>
      </div>
      <div className="drop-title">Déposez un ou plusieurs fichiers ici</div>
      <div className="drop-sub">PDF ou DOCX · baux et avenants acceptés</div>
    </div>
  )
}

function FileQueue({ files, statuses, onRemove }) {
  if (!files.length) return null
  return (
    <div className="file-queue">
      {files.map((f, i) => {
        const st = statuses[i] || {}
        return (
          <div key={i} className={`queue-item ${st.state || 'pending'}`}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14 2 14 8 20 8"/>
            </svg>
            <span className="queue-name">{f.name}</span>
            <span className="queue-size">({(f.size / 1024).toFixed(0)} Ko)</span>
            {st.state === 'loading' && <span className="queue-status">En cours…</span>}
            {st.state === 'done'    && <span className="queue-status success">✓ Extrait</span>}
            {st.state === 'error'   && <span className="queue-status error" title={st.error}>✕ Erreur</span>}
            {!st.state && <button className="queue-remove" onClick={() => onRemove(i)}>✕</button>}
          </div>
        )
      })}
    </div>
  )
}

function AvenantLinkModal({ suggestion, bails, onConfirm, onSkip }) {
  const [selectedId, setSelectedId] = useState(suggestion?.item?.id || null)

  return (
    <div className="modal-overlay">
      <div className="modal">
        <div className="modal-title">Rattacher cet avenant à un bail</div>
        <div className="modal-sub">
          {suggestion
            ? `Bail détecté automatiquement (score ${Math.round(suggestion.score * 100)}%) — confirmez ou choisissez un autre.`
            : 'Sélectionnez le bail d\'origine de cet avenant.'}
        </div>
        <div className="modal-list">
          {bails.filter(b => b.document_type === 'bail').map(b => (
            <button
              key={b.id}
              className={`modal-bail-item${selectedId === b.id ? ' selected' : ''}`}
              onClick={() => setSelectedId(b.id)}
            >
              <div className="modal-bail-name">{b.data?.immeuble || b.data?.adresse || b.file_name}</div>
              <div className="modal-bail-meta">{b.data?.preneur || '—'} · {formatDate(b.created_at)}</div>
            </button>
          ))}
        </div>
        <div className="modal-actions">
          <button className="btn" onClick={onSkip}>Sans rattachement</button>
          <button className="btn primary" disabled={!selectedId} onClick={() => onConfirm(selectedId)}>
            Confirmer
          </button>
        </div>
      </div>
    </div>
  )
}

function FieldCard({ label, value }) {
  return (
    <div className="field-card">
      <div className="field-label">{label}</div>
      <div className={`field-value${!value ? ' empty' : ''}`}>{value || 'Non renseigné'}</div>
    </div>
  )
}

function ResultsView({ item }) {
  const isAvenant = item.document_type === 'avenant'
  const data = isAvenant ? item.data?.champs_modifies || {} : item.data || {}

  return (
    <div className="results">
      {isAvenant && (
        <div className="avenant-banner">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
          </svg>
          <div>
            <strong>Avenant</strong>
            {item.data?.objet_avenant && <span> — {item.data.objet_avenant}</span>}
          </div>
          {item.data?.date_effet_avenant && (
            <span className="avenant-date">Effet : {item.data.date_effet_avenant}</span>
          )}
        </div>
      )}

      {SECTIONS.map(sec => {
        const visibleFields = isAvenant
          ? sec.fields.filter(f => data[f.key] != null)
          : sec.fields
        if (isAvenant && !visibleFields.length) return null
        return (
          <div key={sec.id} className="section-block">
            <div className="section-title">{sec.label}</div>
            <div className="fields-grid">
              {visibleFields.map(f => (
                <FieldCard key={f.key} label={f.label} value={data[f.key]} />
              ))}
            </div>
          </div>
        )
      })}
    </div>
  )
}

function HistoryPanel({ tree, onSelect, activeId }) {
  if (!tree.length) return <div className="history-empty">Aucune extraction sauvegardée</div>
  return (
    <div className="history-list">
      {tree.map(bail => (
        <div key={bail.id}>
          <button
            className={`history-item${bail.id === activeId ? ' active' : ''}`}
            onClick={() => onSelect(bail)}
          >
            <div className="history-name">{bail.data?.immeuble || bail.data?.adresse || bail.file_name}</div>
            <div className="history-meta">{bail.data?.preneur || '—'} · {formatDate(bail.created_at)}</div>
          </button>
          {bail.avenants?.map(av => (
            <button
              key={av.id}
              className={`history-item avenant-item${av.id === activeId ? ' active' : ''}`}
              onClick={() => onSelect(av)}
            >
              <div className="history-name">
                <span className="avenant-tag">A</span>
                {av.data?.objet_avenant || av.file_name}
              </div>
              <div className="history-meta">{formatDate(av.created_at)}</div>
            </button>
          ))}
        </div>
      ))}
      {/* Avenants sans rattachement */}
      {tree.filter(i => i.document_type === 'avenant').map(av => (
        <button
          key={av.id}
          className={`history-item avenant-item${av.id === activeId ? ' active' : ''}`}
          onClick={() => onSelect(av)}
        >
          <div className="history-name">
            <span className="avenant-tag">A</span>
            {av.data?.objet_avenant || av.file_name}
          </div>
          <div className="history-meta">Non rattaché · {formatDate(av.created_at)}</div>
        </button>
      ))}
    </div>
  )
}

// ─── main ────────────────────────────────────────────────────────────────────

export default function App() {
  const [files,         setFiles]         = useState([])
  const [statuses,      setStatuses]      = useState([])
  const [loading,       setLoading]       = useState(false)
  const [activeItem,    setActiveItem]    = useState(null)
  const [history,       setHistory]       = useState([])
  const [histLoaded,    setHistLoaded]    = useState(false)
  const [tab,           setTab]           = useState('extract')
  const [avenantModal,  setAvenantModal]  = useState(null) // { index, extracted, suggestion }

  // ── history tree ──
  function buildTree(rows) {
    const bails   = rows.filter(r => r.document_type === 'bail')
    const avenants = rows.filter(r => r.document_type === 'avenant' && r.parent_id)
    const orphans  = rows.filter(r => r.document_type === 'avenant' && !r.parent_id)
    const tree = bails.map(b => ({
      ...b,
      avenants: avenants.filter(a => a.parent_id === b.id)
    }))
    return [...tree, ...orphans]
  }

  async function loadHistory() {
    if (histLoaded) return
    const { data: rows } = await supabase
      .from('extractions')
      .select('id, file_name, created_at, data, document_type, parent_id')
      .order('created_at', { ascending: false })
      .limit(100)
    if (rows) { setHistory(buildTree(rows)); setHistLoaded(true) }
  }

  function switchTab(t) { setTab(t); if (t === 'history') loadHistory() }

  function setStatus(i, state, error) {
    setStatuses(prev => { const n = [...prev]; n[i] = { state, error }; return n })
  }

  // ── extract one file ──
  async function extractOne(file, index) {
    setStatus(index, 'loading')
    const base64    = await toBase64(file)
    const mediaType = getMediaType(file)

    // First pass : detect if bail or avenant
    const detectPrompt = `Ce document est-il un bail original ou un avenant à un bail existant ? Réponds UNIQUEMENT par le JSON suivant sans markdown : {"type": "bail"} ou {"type": "avenant"}`
    const { data: d1 } = await supabase.functions.invoke('extract-lease', {
      body: { base64, mediaType, prompt: detectPrompt }
    })
    if (d1?.error) throw new Error(d1.error)

    let docType = 'bail'
    try {
      const raw = d1.result
      const parsed = typeof raw === 'string' ? JSON.parse(raw) : raw
      docType = parsed?.type === 'avenant' ? 'avenant' : 'bail'
    } catch (_) {}

    const prompt = docType === 'avenant' ? AVENANT_PROMPT : EXTRACTION_PROMPT
    const { data: d2, error: fnErr } = await supabase.functions.invoke('extract-lease', {
      body: { base64, mediaType, prompt }
    })
    if (fnErr) throw new Error(fnErr.message)
    if (d2?.error) throw new Error(d2.error)

    return { extracted: d2.result, docType }
  }

  // ── save to supabase ──
  async function saveExtraction(file, extracted, docType, parentId) {
    const { data: saved } = await supabase
      .from('extractions')
      .insert({ file_name: file.name, data: extracted, document_type: docType, parent_id: parentId || null })
      .select()
      .single()
    return saved
  }

  // ── main extract handler ──
  async function handleExtract() {
    if (!files.length || loading) return
    setLoading(true)
    setStatuses(files.map(() => ({})))

    const allBails = history.flatMap(b => b.document_type === 'bail' ? [b] : [])

    for (let i = 0; i < files.length; i++) {
      try {
        const { extracted, docType } = await extractOne(files[i], i)

        if (docType === 'avenant') {
          // detect best bail match
          const ref = extracted?.bail_reference
          const match = findBestMatch(ref, allBails)

          // pause and show modal
          await new Promise(resolve => {
            setAvenantModal({ index: i, file: files[i], extracted, suggestion: match, resolve })
          })
        } else {
          const saved = await saveExtraction(files[i], extracted, 'bail', null)
          if (saved) {
            setActiveItem(saved)
            setHistory(prev => {
              const newItem = { ...saved, avenants: [] }
              return [newItem, ...prev]
            })
          }
          setStatus(i, 'done')
        }
      } catch (e) {
        setStatus(i, 'error', e.message)
      }
    }

    setLoading(false)
  }

  async function handleAvenantConfirm(parentId) {
    const { index, file, extracted, resolve } = avenantModal
    setAvenantModal(null)
    try {
      const saved = await saveExtraction(file, extracted, 'avenant', parentId)
      if (saved) {
        setActiveItem(saved)
        setHistory(prev => {
          if (!parentId) return [saved, ...prev]
          return prev.map(b => {
            if (b.id === parentId) return { ...b, avenants: [...(b.avenants || []), saved] }
            return b
          })
        })
      }
      setStatus(index, 'done')
    } catch (e) {
      setStatus(index, 'error', e.message)
    }
    resolve()
  }

  async function handleAvenantSkip() {
    const { index, file, extracted, resolve } = avenantModal
    setAvenantModal(null)
    try {
      const saved = await saveExtraction(file, extracted, 'avenant', null)
      if (saved) { setActiveItem(saved); setHistory(prev => [saved, ...prev]) }
      setStatus(index, 'done')
    } catch (e) {
      setStatus(index, 'error', e.message)
    }
    resolve()
  }

  function handleClear() { setFiles([]); setStatuses([]); setActiveItem(null) }

  const historyBails = history.filter(h => h.document_type === 'bail')

  return (
    <div className="app">
      {avenantModal && (
        <AvenantLinkModal
          suggestion={avenantModal.suggestion}
          bails={historyBails}
          onConfirm={handleAvenantConfirm}
          onSkip={handleAvenantSkip}
        />
      )}

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
          <button className={`nav-item${tab === 'extract' ? ' active' : ''}`} onClick={() => switchTab('extract')}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
            </svg>
            Extraire
          </button>
          <button className={`nav-item${tab === 'history' ? ' active' : ''}`} onClick={() => switchTab('history')}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <polyline points="12 8 12 12 14 14"/>
              <path d="M3.05 11a9 9 0 1 0 .5-4"/><polyline points="3 3 3 7 7 7"/>
            </svg>
            Historique
            {history.length > 0 && <span className="badge">{history.length}</span>}
          </button>
        </nav>

        {tab === 'history' && (
          <HistoryPanel
            tree={history}
            onSelect={item => { setActiveItem(item); setTab('extract') }}
            activeId={activeItem?.id}
          />
        )}
      </aside>

      <main className="main">
        <header className="topbar">
          <h1 className="page-title">
            {activeItem
              ? (activeItem.data?.immeuble || activeItem.data?.adresse || activeItem.file_name)
              : 'Extraction de bail'}
          </h1>
          {activeItem && (
            <div className="topbar-actions">
              <button className="btn" onClick={() => exportToExcel(
                activeItem.document_type === 'avenant' ? activeItem.data?.champs_modifies || {} : activeItem.data || {},
                activeItem.file_name
              )}>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                  <polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                </svg>
                Excel
              </button>
              <button className="btn danger-outline" onClick={handleClear}>Nouveau</button>
            </div>
          )}
        </header>

        <div className="content">
          {!activeItem && (
            <>
              <DropZone onFiles={setFiles} disabled={loading} multi />
              <FileQueue files={files} statuses={statuses} onRemove={i => {
                setFiles(prev => prev.filter((_, j) => j !== i))
                setStatuses(prev => prev.filter((_, j) => j !== i))
              }} />
              {files.length > 0 && !loading && (
                <div className="extract-bar">
                  <button className="btn primary" onClick={handleExtract}>
                    Extraire {files.length > 1 ? `les ${files.length} fichiers` : 'le fichier'}
                  </button>
                  <button className="btn" onClick={handleClear}>Tout effacer</button>
                </div>
              )}
              {loading && (
                <div className="loading-block">
                  <div className="progress-track"><div className="progress-bar active" /></div>
                  <div className="status-msg">Extraction en cours…</div>
                </div>
              )}
            </>
          )}
          {activeItem && <ResultsView item={activeItem} />}
        </div>
      </main>
    </div>
  )
}
