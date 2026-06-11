import { useState, useCallback, useRef } from 'react'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'

// ─── Supabase ────────────────────────────────────────────────────────────────

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
)

// ─── Champs et prompts ───────────────────────────────────────────────────────

const SECTIONS = [
  {
    id: 'actif', label: 'Actif',
    fields: [
      { key: 'adresse',  label: 'Adresse' },
      { key: 'immeuble', label: "Nom de l'immeuble" },
      { key: 'ville',    label: 'Ville / Code postal' },
    ]
  },
  {
    id: 'contrat', label: 'Contrat et durée',
    fields: [
      { key: 'type_bail',    label: 'Type de contrat' },
      { key: 'duree_totale', label: 'Durée totale' },
      { key: 'duree_ferme',  label: 'Durée ferme' },
    ]
  },
  {
    id: 'parties', label: 'Parties',
    fields: [
      { key: 'preneur',  label: 'Preneur' },
      { key: 'bailleur', label: 'Bailleur' },
      { key: 'garant',   label: 'Garant / Caution' },
    ]
  },
  {
    id: 'dates', label: 'Dates clés',
    fields: [
      { key: 'date_effet',          label: "Date d'effet" },
      { key: 'date_signature',      label: 'Date de signature' },
      { key: 'break_option',        label: 'Break option' },
      { key: 'notice',              label: 'Préavis (notice)' },
      { key: 'date_conge',          label: 'Date limite de congé' },
      { key: 'date_fin',            label: 'Date de fin' },
      { key: 'date_limite_travaux', label: 'Date limite travaux preneur' },
      { key: 'conditions_break',    label: 'Conditions financières du break' },
    ]
  },
  {
    id: 'surfaces', label: 'Surfaces',
    fields: [
      { key: 'surface_bureaux',  label: 'Surface bureaux (m²)' },
      { key: 'surface_totale',   label: 'Surface totale (m²)' },
      { key: 'parking',          label: 'Parking' },
      { key: 'rie',              label: 'RIE' },
      { key: 'autres_surfaces',  label: 'Autres surfaces' },
    ]
  },
  {
    id: 'loyer', label: 'Loyer, taxes et charges',
    fields: [
      { key: 'loyer_signature',         label: 'Loyer HT/HC à la signature (€)' },
      { key: 'loyer_cours',             label: 'Loyer HT/HC en cours (€)' },
      { key: 'indexation',              label: 'Indexation / indice' },
      { key: 'franchise',               label: 'Franchise' },
      { key: 'participation_travaux',   label: 'Participation travaux bailleur (€)' },
      { key: 'travaux_bailleur_preneur',label: 'Détail travaux financés par le bailleur' },
      { key: 'indemnite_depart',        label: 'Indemnités en cas de départ du preneur' },
      { key: 'depot_garantie',          label: 'Dépôt de garantie' },
      { key: 'charges',                 label: 'Charges / TEOM' },
    ]
  },
  {
    id: 'jouissance', label: 'Refacturation et jouissance',
    fields: [
      { key: 'article_606',    label: 'Article 606' },
      { key: 'conformite',     label: 'Conformité' },
      { key: 'accession',      label: 'Accession' },
      { key: 'remise_en_etat', label: 'Remise en état' },
      { key: 'maintenance',    label: 'Maintenance & travaux' },
      { key: 'destination',    label: 'Destination' },
      { key: 'sous_location',  label: 'Sous-location' },
      { key: 'cession',        label: 'Cession' },
    ]
  }
]

const ALL_FIELDS = SECTIONS.flatMap(s => s.fields)

const EXTRACTION_PROMPT = `Tu es un expert en baux commerciaux français. Analyse ce document et extrait précisément les données suivantes. Réponds UNIQUEMENT en JSON valide, sans markdown ni backticks, avec exactement ces clés :

{
  "adresse": "adresse complète",
  "immeuble": "nom de l'immeuble si mentionné",
  "ville": "ville et code postal",
  "type_bail": "type de bail (commercial, dérogatoire, convention d'occupation, etc.)",
  "duree_totale": "durée totale",
  "duree_ferme": "durée ferme",
  "preneur": "nom du preneur / locataire",
  "bailleur": "nom du bailleur",
  "garant": "garant ou caution si mentionné",
  "date_effet": "date d'effet / prise d'effet",
  "date_signature": "date de signature",
  "break_option": "date de break option / résiliation triennale",
  "notice": "durée de préavis",
  "date_conge": "date limite pour donner congé",
  "date_fin": "date d'expiration du bail",
  "date_limite_travaux": "date limite pour réaliser les travaux preneur, condition d'octroi de la participation bailleur",
  "conditions_break": "conditions financières et formelles d'exercice du break : préavis, indemnités éventuelles, état des lieux anticipé",
  "surface_bureaux": "détail des surfaces bureaux par niveau/bâtiment",
  "surface_totale": "surface totale en m²",
  "parking": "nombre et description des places de parking",
  "rie": "restaurant inter-entreprises : oui/non et modalités",
  "autres_surfaces": "autres surfaces (archives, locaux techniques, etc.)",
  "loyer_signature": "loyer annuel HT/HC à la signature",
  "loyer_cours": "loyer annuel HT/HC actuel",
  "indexation": "clause d'indexation et indice (ILC, ILAT, ICC...)",
  "franchise": "franchise de loyer (durée, modalités)",
  "participation_travaux": "montant de la participation bailleur aux travaux preneur",
  "travaux_bailleur_preneur": "détail des travaux financés par le bailleur : montant, conditions d'appel, sort en cas de non-consommation, date limite",
  "indemnite_depart": "indemnités dues par le preneur en cas de départ : restitution de franchise, indemnité libératoire de remise en état, autres pénalités",
  "depot_garantie": "montant et durée du dépôt de garantie",
  "charges": "répartition des charges, TEOM, provisions",
  "article_606": "qui supporte l'article 606",
  "conformite": "obligations de conformité",
  "accession": "clause d'accession",
  "remise_en_etat": "obligations de remise en état",
  "maintenance": "obligations de maintenance et travaux en cours de bail",
  "destination": "destination contractuelle des locaux",
  "sous_location": "conditions de sous-location",
  "cession": "conditions de cession du bail"
}

Si une information est absente du document, mets null. Reprends les montants, dates et formulations exactes du document.`

const AVENANT_PROMPT = `Tu es un expert en baux commerciaux français. Ce document est un AVENANT à un bail existant.

Réponds UNIQUEMENT en JSON valide, sans markdown ni backticks, avec exactement cette structure :

{
  "bail_reference": {
    "preneur": "nom du preneur mentionné dans l'avenant",
    "bailleur": "nom du bailleur mentionné dans l'avenant",
    "date_bail_origine": "date du bail d'origine mentionnée dans l'avenant",
    "adresse": "adresse de l'immeuble mentionnée dans l'avenant",
    "immeuble": "nom de l'immeuble si mentionné"
  },
  "date_effet_avenant": "date d'entrée en vigueur de l'avenant",
  "date_signature_avenant": "date de signature de l'avenant",
  "objet_avenant": "résumé en 1-2 phrases de l'objet principal de l'avenant",
  "champs_modifies": {
    "adresse": null, "immeuble": null, "ville": null, "type_bail": null,
    "duree_totale": null, "duree_ferme": null, "preneur": null, "bailleur": null,
    "garant": null, "date_effet": null, "date_signature": null, "break_option": null,
    "notice": null, "date_conge": null, "date_fin": null, "date_limite_travaux": null,
    "conditions_break": null, "surface_bureaux": null, "surface_totale": null,
    "parking": null, "rie": null, "autres_surfaces": null, "loyer_signature": null,
    "loyer_cours": null, "indexation": null, "franchise": null,
    "participation_travaux": null, "travaux_bailleur_preneur": null,
    "indemnite_depart": null, "depot_garantie": null, "charges": null,
    "article_606": null, "conformite": null, "accession": null,
    "remise_en_etat": null, "maintenance": null, "destination": null,
    "sous_location": null, "cession": null
  }
}

IMPORTANT : dans "champs_modifies", ne renseigne QUE les champs effectivement modifiés par cet avenant. Laisse null tous les autres. Reprends les montants, dates et formulations exactes du document.`

// ─── CSS ─────────────────────────────────────────────────────────────────────

const CSS = `
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --bg: #f8f7f5; --surface: #ffffff; --surface2: #f3f2ef;
  --border: rgba(0,0,0,0.10); --border2: rgba(0,0,0,0.18);
  --text: #1a1917; --text2: #5a5855; --text3: #9a9895;
  --accent: #185FA5; --accent-light: #E6F1FB; --accent-dark: #0C447C;
  --danger: #A32D2D; --danger-light: #FCEBEB;
  --success: #0F6E56; --success-light: #E1F5EE;
  --r: 8px; --rl: 12px;
  font-family: 'Inter', system-ui, -apple-system, sans-serif;
}
body { background: var(--bg); color: var(--text); font-size: 14px; line-height: 1.6; min-height: 100vh; }
button, input, select { font-family: inherit; font-size: 13px; cursor: pointer; }
::-webkit-scrollbar { width: 6px; } ::-webkit-scrollbar-thumb { background: var(--border2); border-radius: 3px; }

.app { display: flex; min-height: 100vh; }

.sidebar { width: 220px; flex-shrink: 0; background: var(--surface); border-right: 1px solid var(--border); display: flex; flex-direction: column; }
.sidebar-logo { display: flex; align-items: center; gap: 10px; padding: 20px 16px 16px; font-size: 14px; font-weight: 600; color: var(--accent); border-bottom: 1px solid var(--border); }
.sidebar-nav { padding: 12px 8px; display: flex; flex-direction: column; gap: 2px; }
.nav-item { display: flex; align-items: center; gap: 9px; padding: 8px 10px; border-radius: var(--r); border: none; background: transparent; color: var(--text2); font-size: 13px; font-weight: 500; text-align: left; width: 100%; transition: background .15s, color .15s; }
.nav-item:hover { background: var(--surface2); color: var(--text); }
.nav-item.active { background: var(--accent-light); color: var(--accent); }
.nav-item .badge { margin-left: auto; background: var(--accent-light); color: var(--accent); font-size: 11px; font-weight: 600; padding: 1px 6px; border-radius: 999px; }
.sidebar-footer { padding: 8px 12px 12px; border-top: 1px solid var(--border); margin-top: auto; }
.btn-clear-history { width: 100%; padding: 7px 10px; border-radius: var(--r); border: 1px solid #F09595; background: transparent; color: var(--danger); font-size: 12px; font-weight: 500; transition: background .15s; }
.btn-clear-history:hover { background: var(--danger-light); }

.history-list { flex: 1; overflow-y: auto; padding: 0 8px 12px; display: flex; flex-direction: column; gap: 2px; }
.history-empty { padding: 16px; font-size: 12px; color: var(--text3); text-align: center; }
.history-item { display: block; width: 100%; text-align: left; padding: 9px 10px; border: none; border-radius: var(--r); background: transparent; transition: background .15s; }
.history-item:hover { background: var(--surface2); }
.history-item.active { background: var(--accent-light); }
.history-name { font-size: 12px; font-weight: 500; color: var(--text); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.history-meta { font-size: 11px; color: var(--text3); margin-top: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.avenant-item { padding-left: 24px !important; border-left: 2px solid #B5D4F4 !important; margin-left: 8px; }
.avenant-tag { display: inline-flex; align-items: center; justify-content: center; width: 16px; height: 16px; border-radius: 3px; background: var(--accent); color: #fff; font-size: 10px; font-weight: 700; margin-right: 6px; flex-shrink: 0; }

.main { flex: 1; display: flex; flex-direction: column; min-width: 0; }
.topbar { display: flex; align-items: center; padding: 16px 28px; border-bottom: 1px solid var(--border); background: var(--surface); gap: 16px; flex-shrink: 0; }
.page-title { font-size: 15px; font-weight: 600; color: var(--text); flex: 1; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.topbar-actions { display: flex; gap: 8px; }
.content { flex: 1; padding: 28px; overflow-y: auto; max-width: 960px; width: 100%; }

.drop-zone { border: 1.5px dashed var(--border2); border-radius: var(--rl); padding: 48px 24px; text-align: center; cursor: pointer; transition: border-color .2s, background .2s; background: var(--surface); user-select: none; }
.drop-zone:hover, .drop-zone.dragging { border-color: var(--accent); background: var(--accent-light); }
.drop-zone.disabled { opacity: .5; cursor: not-allowed; pointer-events: none; }
.drop-icon { color: var(--text3); margin-bottom: 12px; }
.drop-title { font-size: 14px; font-weight: 500; color: var(--text); margin-bottom: 4px; }
.drop-sub { font-size: 12px; color: var(--text3); }

.file-queue { display: flex; flex-direction: column; gap: 6px; margin-top: 12px; }
.queue-item { display: flex; align-items: center; gap: 8px; padding: 9px 14px; background: var(--surface); border: 1px solid var(--border); border-radius: var(--r); color: var(--text2); font-size: 13px; }
.queue-item.loading { border-color: #B5D4F4; background: var(--accent-light); }
.queue-item.done    { border-color: #9FE1CB; background: var(--success-light); }
.queue-item.error   { border-color: #F09595; background: var(--danger-light); }
.queue-name { font-weight: 500; color: var(--text); flex: 1; }
.queue-size { font-size: 12px; color: var(--text3); }
.queue-status { font-size: 12px; margin-left: auto; }
.queue-status.success { color: var(--success); }
.queue-status.error   { color: var(--danger); cursor: help; }
.queue-remove { margin-left: auto; background: none; border: none; color: var(--text3); font-size: 14px; padding: 0 2px; }
.queue-remove:hover { color: var(--danger); }
.extract-bar { display: flex; gap: 8px; margin-top: 16px; }

.btn { display: inline-flex; align-items: center; gap: 6px; padding: 7px 13px; border-radius: var(--r); border: 1px solid var(--border2); background: var(--surface); color: var(--text); font-weight: 500; font-size: 13px; transition: background .15s; white-space: nowrap; }
.btn:hover { background: var(--surface2); }
.btn.primary { background: var(--accent); color: #fff; border-color: var(--accent); }
.btn.primary:hover { background: var(--accent-dark); border-color: var(--accent-dark); }
.btn.danger-outline { border-color: var(--danger); color: var(--danger); }
.btn.danger-outline:hover { background: var(--danger-light); }
.btn:disabled { opacity: .4; cursor: not-allowed; }

.loading-block { margin-top: 16px; }
.progress-track { height: 3px; background: var(--surface2); border-radius: 2px; overflow: hidden; margin-bottom: 8px; }
.progress-bar { height: 100%; width: 0; background: var(--accent); border-radius: 2px; }
.progress-bar.active { animation: prog 8s ease-out forwards; }
@keyframes prog { 0%{width:0%} 60%{width:70%} 100%{width:88%} }
.status-msg { font-size: 12px; color: var(--text3); }

.modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.4); display: flex; align-items: center; justify-content: center; z-index: 100; }
.modal { background: var(--surface); border-radius: var(--rl); border: 1px solid var(--border); padding: 24px; width: 480px; max-width: 94vw; max-height: 80vh; display: flex; flex-direction: column; gap: 12px; }
.modal-title { font-size: 15px; font-weight: 600; color: var(--text); }
.modal-sub   { font-size: 13px; color: var(--text2); }
.modal-list  { display: flex; flex-direction: column; gap: 4px; overflow-y: auto; max-height: 280px; }
.modal-bail-item { display: block; width: 100%; text-align: left; padding: 10px 12px; border-radius: var(--r); border: 1px solid var(--border); background: var(--surface); cursor: pointer; transition: background .15s; }
.modal-bail-item:hover { background: var(--surface2); }
.modal-bail-item.selected { border-color: var(--accent); background: var(--accent-light); }
.modal-bail-name { font-size: 13px; font-weight: 500; color: var(--text); }
.modal-bail-meta { font-size: 11px; color: var(--text3); margin-top: 2px; }
.modal-actions { display: flex; gap: 8px; justify-content: flex-end; padding-top: 4px; }

.avenant-banner { display: flex; align-items: flex-start; gap: 10px; background: var(--accent-light); border: 1px solid #B5D4F4; border-radius: var(--r); padding: 12px 14px; font-size: 13px; color: var(--accent-dark); margin-bottom: 20px; }
.avenant-date { margin-left: auto; font-size: 12px; white-space: nowrap; }

.results { display: flex; flex-direction: column; }
.section-block { margin-bottom: 28px; }
.section-title { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .07em; color: var(--text3); margin-bottom: 10px; padding-bottom: 6px; border-bottom: 1px solid var(--border); }
.fields-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: 8px; }
.field-card { background: var(--surface); border: 1px solid var(--border); border-radius: var(--r); padding: 10px 14px; }
.field-label { font-size: 11px; color: var(--text3); margin-bottom: 3px; }
.field-value { font-size: 13px; color: var(--text); font-weight: 500; line-height: 1.5; }
.field-value.empty { color: var(--text3); font-style: italic; font-weight: 400; }
`

// ─── helpers ─────────────────────────────────────────────────────────────────

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
  const na = a.toLowerCase().trim(), nb = b.toLowerCase().trim()
  if (na === nb) return 1
  if (na.includes(nb) || nb.includes(na)) return 0.8
  const wa = na.split(/\s+/), wb = nb.split(/\s+/)
  return wa.filter(w => wb.includes(w)).length / Math.max(wa.length, wb.length)
}

function findBestMatch(ref, bails) {
  if (!ref || !bails.length) return null
  let best = null, bestScore = 0
  for (const b of bails) {
    const d = b.data || {}
    const score = similarity(ref.preneur, d.preneur) * 0.4 + similarity(ref.bailleur, d.bailleur) * 0.2
      + similarity(ref.adresse, d.adresse) * 0.2 + similarity(ref.immeuble, d.immeuble) * 0.2
    if (score > bestScore) { bestScore = score; best = b }
  }
  return bestScore > 0.3 ? { item: best, score: bestScore } : null
}

function exportToExcel(data, fileName) {
  const wb = XLSX.utils.book_new()
  const headers = ALL_FIELDS.map(f => f.label)
  const values  = ALL_FIELDS.map(f => data[f.key] ?? '')
  const ws1 = XLSX.utils.aoa_to_sheet([headers, values])
  ws1['!cols'] = headers.map(() => ({ wch: 24 }))
  XLSX.utils.book_append_sheet(wb, ws1, 'Tableau')
  const rows = [['Section', 'Champ', 'Valeur'], ...SECTIONS.flatMap(s => s.fields.map(f => [s.label, f.label, data[f.key] ?? '']))]
  const ws2 = XLSX.utils.aoa_to_sheet(rows)
  ws2['!cols'] = [{ wch: 26 }, { wch: 32 }, { wch: 64 }]
  XLSX.utils.book_append_sheet(wb, ws2, 'Fiche')
  XLSX.writeFile(wb, `lease_abstract_${(fileName || 'bail').replace(/\.[^.]+$/, '')}.xlsx`)
}

// ─── sub-components ───────────────────────────────────────────────────────────

function DropZone({ onFiles, disabled }) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef()
  const handle = useCallback(files => {
    const valid = Array.from(files).filter(f => ['pdf','docx'].includes(f.name.split('.').pop().toLowerCase()))
    if (valid.length) onFiles(valid)
    else alert('Format non supporté. Utilisez des fichiers PDF ou DOCX.')
  }, [onFiles])
  return (
    <div className={`drop-zone${dragging?' dragging':''}${disabled?' disabled':''}`}
      onClick={() => !disabled && inputRef.current?.click()}
      onDragOver={e => { e.preventDefault(); setDragging(true) }}
      onDragLeave={() => setDragging(false)}
      onDrop={e => { e.preventDefault(); setDragging(false); handle(e.dataTransfer.files) }}>
      <input ref={inputRef} type="file" accept=".pdf,.docx" multiple style={{ display:'none' }} onChange={e => handle(e.target.files)} />
      <div className="drop-icon">
        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
          <polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>
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
          <div key={i} className={`queue-item ${st.state||'pending'}`}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
            </svg>
            <span className="queue-name">{f.name}</span>
            <span className="queue-size">({(f.size/1024).toFixed(0)} Ko)</span>
            {st.state==='loading' && <span className="queue-status">En cours…</span>}
            {st.state==='done'    && <span className="queue-status success">✓ Extrait</span>}
            {st.state==='error'   && <span className="queue-status error" title={st.error}>✕ Erreur</span>}
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
            ? `Bail détecté automatiquement (score ${Math.round(suggestion.score*100)}%) — confirmez ou choisissez un autre.`
            : "Sélectionnez le bail d'origine de cet avenant."}
        </div>
        <div className="modal-list">
          {bails.map(b => (
            <button key={b.id} className={`modal-bail-item${selectedId===b.id?' selected':''}`} onClick={() => setSelectedId(b.id)}>
              <div className="modal-bail-name">{b.data?.immeuble || b.data?.adresse || b.file_name}</div>
              <div className="modal-bail-meta">{b.data?.preneur||'—'} · {formatDate(b.created_at)}</div>
            </button>
          ))}
        </div>
        <div className="modal-actions">
          <button className="btn" onClick={onSkip}>Sans rattachement</button>
          <button className="btn primary" disabled={!selectedId} onClick={() => onConfirm(selectedId)}>Confirmer</button>
        </div>
      </div>
    </div>
  )
}

function FieldCard({ label, value }) {
  return (
    <div className="field-card">
      <div className="field-label">{label}</div>
      <div className={`field-value${!value?' empty':''}`}>{value||'Non renseigné'}</div>
    </div>
  )
}

function ResultsView({ item }) {
  const isAvenant = item.document_type === 'avenant'
  const data = isAvenant ? item.data?.champs_modifies||{} : item.data||{}
  return (
    <div className="results">
      {isAvenant && (
        <div className="avenant-banner">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
          </svg>
          <div><strong>Avenant</strong>{item.data?.objet_avenant && <span> — {item.data.objet_avenant}</span>}</div>
          {item.data?.date_effet_avenant && <span className="avenant-date">Effet : {item.data.date_effet_avenant}</span>}
        </div>
      )}
      {SECTIONS.map(sec => {
        const fields = isAvenant ? sec.fields.filter(f => data[f.key]!=null) : sec.fields
        if (isAvenant && !fields.length) return null
        return (
          <div key={sec.id} className="section-block">
            <div className="section-title">{sec.label}</div>
            <div className="fields-grid">
              {fields.map(f => <FieldCard key={f.key} label={f.label} value={data[f.key]} />)}
            </div>
          </div>
        )
      })}
    </div>
  )
}

function HistoryPanel({ tree, onSelect, activeId, onClear }) {
  return (
    <>
      <div className="history-list">
        {!tree.length
          ? <div className="history-empty">Aucune extraction sauvegardée</div>
          : tree.map(bail => (
            <div key={bail.id}>
              <button className={`history-item${bail.id===activeId?' active':''}`} onClick={() => onSelect(bail)}>
                <div className="history-name">{bail.data?.immeuble||bail.data?.adresse||bail.file_name}</div>
                <div className="history-meta">{bail.data?.preneur||'—'} · {formatDate(bail.created_at)}</div>
              </button>
              {bail.avenants?.map(av => (
                <button key={av.id} className={`history-item avenant-item${av.id===activeId?' active':''}`} onClick={() => onSelect(av)}>
                  <div className="history-name"><span className="avenant-tag">A</span>{av.data?.objet_avenant||av.file_name}</div>
                  <div className="history-meta">{formatDate(av.created_at)}</div>
                </button>
              ))}
            </div>
          ))
        }
      </div>
      {tree.length > 0 && (
        <div className="sidebar-footer">
          <button className="btn-clear-history" onClick={onClear}>Vider l'historique</button>
        </div>
      )}
    </>
  )
}

// ─── main ─────────────────────────────────────────────────────────────────────

export default function App() {
  const [files,        setFiles]        = useState([])
  const [statuses,     setStatuses]     = useState([])
  const [loading,      setLoading]      = useState(false)
  const [activeItem,   setActiveItem]   = useState(null)
  const [history,      setHistory]      = useState([])
  const [histLoaded,   setHistLoaded]   = useState(false)
  const [tab,          setTab]          = useState('extract')
  const [avenantModal, setAvenantModal] = useState(null)

  function buildTree(rows) {
    const bails    = rows.filter(r => r.document_type === 'bail')
    const avenants = rows.filter(r => r.document_type === 'avenant' && r.parent_id)
    const orphans  = rows.filter(r => r.document_type === 'avenant' && !r.parent_id)
    return [...bails.map(b => ({ ...b, avenants: avenants.filter(a => a.parent_id === b.id) })), ...orphans]
  }

  async function loadHistory() {
    if (histLoaded) return
    const { data: rows } = await supabase.from('extractions')
      .select('id, file_name, created_at, data, document_type, parent_id')
      .order('created_at', { ascending: false }).limit(100)
    if (rows) { setHistory(buildTree(rows)); setHistLoaded(true) }
  }

  function switchTab(t) { setTab(t); if (t === 'history') loadHistory() }

  function setStatus(i, state, error) {
    setStatuses(prev => { const n=[...prev]; n[i]={state,error}; return n })
  }

  async function extractOne(file, index) {
    setStatus(index, 'loading')
    const base64 = await toBase64(file)
    const mediaType = getMediaType(file)
    const detectPrompt = `Ce document est-il un bail original ou un avenant ? Réponds UNIQUEMENT par le JSON suivant sans markdown : {"type": "bail"} ou {"type": "avenant"}`
    const { data: d1 } = await supabase.functions.invoke('extract-lease', { body: { base64, mediaType, prompt: detectPrompt } })
    if (d1?.error) throw new Error(d1.error)
    let docType = 'bail'
    try { const p = typeof d1.result==='string' ? JSON.parse(d1.result) : d1.result; docType = p?.type==='avenant' ? 'avenant' : 'bail' } catch(_) {}
    const { data: d2, error: fnErr } = await supabase.functions.invoke('extract-lease', { body: { base64, mediaType, prompt: docType==='avenant' ? AVENANT_PROMPT : EXTRACTION_PROMPT } })
    if (fnErr) throw new Error(fnErr.message)
    if (d2?.error) throw new Error(d2.error)
    return { extracted: d2.result, docType }
  }

  async function saveExtraction(file, extracted, docType, parentId) {
    const { data: saved } = await supabase.from('extractions')
      .insert({ file_name: file.name, data: extracted, document_type: docType, parent_id: parentId||null })
      .select().single()
    return saved
  }

  async function handleExtract() {
    if (!files.length || loading) return
    setLoading(true)
    setStatuses(files.map(() => ({})))
    const allBails = history.filter(h => h.document_type === 'bail')
    for (let i = 0; i < files.length; i++) {
      try {
        const { extracted, docType } = await extractOne(files[i], i)
        if (docType === 'avenant') {
          const match = findBestMatch(extracted?.bail_reference, allBails)
          await new Promise(resolve => setAvenantModal({ index: i, file: files[i], extracted, suggestion: match, resolve }))
        } else {
          const saved = await saveExtraction(files[i], extracted, 'bail', null)
          if (saved) { setActiveItem(saved); setHistory(prev => [{ ...saved, avenants: [] }, ...prev]) }
          setStatus(i, 'done')
        }
      } catch(e) { setStatus(i, 'error', e.message) }
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
        setHistory(prev => parentId
          ? prev.map(b => b.id===parentId ? { ...b, avenants: [...(b.avenants||[]), saved] } : b)
          : [saved, ...prev])
      }
      setStatus(index, 'done')
    } catch(e) { setStatus(index, 'error', e.message) }
    resolve()
  }

  async function handleAvenantSkip() {
    const { index, file, extracted, resolve } = avenantModal
    setAvenantModal(null)
    try {
      const saved = await saveExtraction(file, extracted, 'avenant', null)
      if (saved) { setActiveItem(saved); setHistory(prev => [saved, ...prev]) }
      setStatus(index, 'done')
    } catch(e) { setStatus(index, 'error', e.message) }
    resolve()
  }

  async function handleClearHistory() {
    if (!window.confirm("Vider tout l'historique ? Cette action est irréversible.")) return
    await supabase.from('extractions').delete().neq('id', '00000000-0000-0000-0000-000000000000')
    setHistory([]); setHistLoaded(false); setActiveItem(null)
  }

  function handleClear() { setFiles([]); setStatuses([]); setActiveItem(null) }

  return (
    <>
      <style>{CSS}</style>
      <div className="app">
        {avenantModal && (
          <AvenantLinkModal
            suggestion={avenantModal.suggestion}
            bails={history.filter(h => h.document_type==='bail')}
            onConfirm={handleAvenantConfirm}
            onSkip={handleAvenantSkip}
          />
        )}

        <aside className="sidebar">
          <div className="sidebar-logo">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14 2 14 8 20 8"/>
              <line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>
            </svg>
            <span>Lease Reader</span>
          </div>
          <nav className="sidebar-nav">
            <button className={`nav-item${tab==='extract'?' active':''}`} onClick={() => switchTab('extract')}>
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
              </svg>
              Extraire
            </button>
            <button className={`nav-item${tab==='history'?' active':''}`} onClick={() => switchTab('history')}>
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
              onClear={handleClearHistory}
            />
          )}
        </aside>

        <main className="main">
          <header className="topbar">
            <h1 className="page-title">
              {activeItem ? (activeItem.data?.immeuble||activeItem.data?.adresse||activeItem.file_name) : 'Extraction de bail'}
            </h1>
            {activeItem && (
              <div className="topbar-actions">
                <button className="btn" onClick={() => exportToExcel(
                  activeItem.document_type==='avenant' ? activeItem.data?.champs_modifies||{} : activeItem.data||{},
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
                <DropZone onFiles={setFiles} disabled={loading} />
                <FileQueue files={files} statuses={statuses} onRemove={i => {
                  setFiles(p => p.filter((_,j)=>j!==i)); setStatuses(p => p.filter((_,j)=>j!==i))
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
    </>
  )
}
