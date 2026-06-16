import { useState, useCallback, useRef } from 'react'
import './index.css'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
)

// ─── Sections ────────────────────────────────────────────────────────────────

const SECTIONS = [
  { id: 'actif', label: 'Actif', fields: [
    { key: 'adresse', label: 'Adresse' },
    { key: 'immeuble', label: "Nom de l'immeuble" },
    { key: 'ville', label: 'Ville / Code postal' },
  ]},
  { id: 'contrat', label: 'Contrat et durée', fields: [
    { key: 'type_bail', label: 'Type de contrat' },
    { key: 'duree_totale', label: 'Durée totale' },
    { key: 'duree_ferme', label: 'Durée ferme' },
  ]},
  { id: 'parties', label: 'Parties', fields: [
    { key: 'preneur', label: 'Preneur' },
    { key: 'bailleur', label: 'Bailleur' },
    { key: 'garant', label: 'Garant / Caution' },
  ]},
  { id: 'dates', label: 'Dates clés', fields: [
    { key: 'date_effet', label: "Date d'effet" },
    { key: 'date_signature', label: 'Date de signature' },
    { key: 'break_options', label: 'Break options' },
    { key: 'notice', label: 'Préavis' },
    { key: 'date_conge', label: 'Date limite de congé' },
    { key: 'date_fin', label: 'Date de fin' },
    { key: 'date_limite_travaux', label: 'Date limite travaux preneur' },
    { key: 'conditions_break', label: 'Conditions financières du break' },
  ]},
  { id: 'surfaces', label: 'Surfaces', fields: [
    { key: 'surface_totale_m2', label: 'Surface totale (m²) — valeur brute' },
    { key: 'surfaces_detail', label: 'Tableau surfaces par typologie' },
    { key: 'surface_bureaux', label: 'Détail surfaces bureaux' },
    { key: 'surface_totale', label: 'Surface totale (formulation complète)' },
    { key: 'parking', label: 'Parking' },
    { key: 'rie', label: 'RIE' },
    { key: 'autres_surfaces', label: 'Autres surfaces' },
  ]},
  { id: 'loyer', label: 'Loyer', fields: [
    { key: 'loyer_signature_montant', label: 'Loyer à la signature — montant brut (€/an)' },
    { key: 'loyer_signature', label: 'Loyer à la signature — formulation complète' },
    { key: 'loyer_cours', label: 'Loyer en cours' },
    { key: 'indexation', label: 'Indexation / indice' },
    { key: 'franchise_duree', label: 'Franchise — durée' },
    { key: 'franchise', label: 'Franchise — modalités complètes' },
    { key: 'charges', label: 'Charges / TEOM' },
  ]},
  { id: 'depot', label: 'Dépôt de garantie', fields: [
    { key: 'depot_garantie_montant', label: 'Dépôt de garantie — montant brut (€)' },
    { key: 'depot_garantie', label: 'Dépôt de garantie — modalités complètes' },
  ]},
  { id: 'travaux', label: 'Participation travaux bailleur', fields: [
    { key: 'travaux_montant', label: 'Montant (€)' },
    { key: 'travaux_date_factures', label: 'Date limite réception factures' },
    { key: 'travaux_modalites', label: 'Modalités complètes' },
  ]},
  { id: 'indemnites', label: 'Indemnités contractuelles', fields: [
    { key: 'indemnites', label: 'Tableau structuré des indemnités' },
    { key: 'indemnites_detail', label: 'Détail verbeux' },
  ]},
  { id: 'jouissance', label: 'Refacturation et jouissance', fields: [
    { key: 'article_606', label: 'Article 606' },
    { key: 'conformite', label: 'Conformité' },
    { key: 'accession', label: 'Accession' },
    { key: 'remise_en_etat', label: 'Remise en état' },
    { key: 'maintenance', label: 'Maintenance & travaux' },
    { key: 'destination', label: 'Destination' },
    { key: 'sous_location', label: 'Sous-location' },
    { key: 'cession', label: 'Cession' },
  ]},
]

const ALL_FIELDS = SECTIONS.flatMap(s => s.fields)

// ─── Prompts ─────────────────────────────────────────────────────────────────

const EXTRACTION_PROMPT = `Expert baux commerciaux français. Extrais les données du document. JSON valide uniquement, sans markdown.

REGLES: Utilise uniquement des guillemets droits ASCII dans le JSON. Remplace tout caractère typographique par son equivalent ASCII. Pas de retour a la ligne dans les valeurs. Echappe les guillemets internes avec backslash. break_options=tableau strings. franchise_periodes=tableau objets. indemnites=tableau objets. Champs _montant=chiffres bruts. null si absent.

Retourne exactement cette structure avec les valeurs du document:
{"adresse":null,"immeuble":null,"ville":null,"type_bail":null,"duree_totale":null,"duree_ferme":null,"preneur":null,"bailleur":null,"garant":null,"date_effet":null,"date_signature":null,"break_options":[],"notice":null,"date_conge":null,"date_fin":null,"date_limite_travaux":null,"conditions_break":null,"surface_totale_m2":null,"surfaces_detail":[],"surface_bureaux":null,"surface_totale":null,"parking":null,"rie":null,"autres_surfaces":null,"loyer_signature_montant":null,"loyer_signature":null,"loyer_cours":null,"indexation":null,"franchise_periodes":[],"franchise":null,"charges":null,"depot_garantie_montant":null,"depot_garantie":null,"travaux_montant":null,"travaux_date_factures":null,"travaux_modalites":null,"indemnites":[],"indemnites_detail":null,"article_606":null,"conformite":null,"accession":null,"remise_en_etat":null,"maintenance":null,"destination":null,"sous_location":null,"cession":null}

Formats: surfaces_detail=[{"typologie":"Bureaux","localisation":"5ème étage","surface_m2":"2224.98","prix_unitaire":"290","loyer_annuel":"645244"}] | break_options=["31/08/2027","31/08/2030"] | franchise_periodes=[{"date_debut":"jj/mm/aa","date_fin":"jj/mm/aa","duree":"3 mois","montant":"123405","indexation_incluse":"Oui/Non/Non précisé","condition":null}] | indemnites=[{"motif":"...","due_par":"Preneur ou Bailleur","montant":"...","date_limite":"..."}]`

const AVENANT_PROMPT = `Expert baux commerciaux français. Ce document est un AVENANT. JSON valide uniquement, sans markdown.

RÈGLES: Ne renseigne dans champs_modifies QUE les champs modifiés. null pour les autres. break_options/franchise_periodes/indemnites=tableaux si modifiés.

{"bail_reference":{"preneur":null,"bailleur":null,"date_bail_origine":null,"adresse":null,"immeuble":null},"date_effet_avenant":null,"date_signature_avenant":null,"objet_avenant":null,"champs_modifies":{"adresse":null,"immeuble":null,"ville":null,"type_bail":null,"duree_totale":null,"duree_ferme":null,"preneur":null,"bailleur":null,"garant":null,"date_effet":null,"date_signature":null,"break_options":null,"notice":null,"date_conge":null,"date_fin":null,"date_limite_travaux":null,"conditions_break":null,"surface_totale_m2":null,"surfaces_detail":[],"surface_bureaux":null,"surface_totale":null,"parking":null,"rie":null,"autres_surfaces":null,"loyer_signature_montant":null,"loyer_signature":null,"loyer_cours":null,"indexation":null,"franchise_periodes":null,"franchise":null,"charges":null,"depot_garantie_montant":null,"depot_garantie":null,"travaux_montant":null,"travaux_date_factures":null,"travaux_modalites":null,"indemnites":null,"indemnites_detail":null,"article_606":null,"conformite":null,"accession":null,"remise_en_etat":null,"maintenance":null,"destination":null,"sous_location":null,"cession":null}}`

// ─── Helpers ─────────────────────────────────────────────────────────────────

function toBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader()
    r.onload  = () => res(r.result.split(',')[1])
    r.onerror = () => rej(new Error('Lecture échouée'))
    r.readAsDataURL(file)
  })
}
function getMediaType(file) {
  return file.name.toLowerCase().endsWith('.pdf') ? 'application/pdf'
    : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
}
function formatDate(iso) {
  if (!iso) return '—'
  return new Date(iso).toLocaleDateString('fr-FR', { day: '2-digit', month: 'short', year: 'numeric' })
}
function fmtEur(val) {
  if (!val) return null
  const n = parseFloat(String(val).replace(/[^0-9.]/g, ''))
  if (isNaN(n)) return val
  return new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 }).format(n)
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

  // Champs simples (hors tableaux structurés)
  const simpleFields = ALL_FIELDS.filter(f => !['break_options','franchise_periodes','indemnites'].includes(f.key))

  // Colonnes dynamiques pour break_options (max 4)
  const MAX_BREAKS = 4
  const breakCols = []
  for (let i = 0; i < MAX_BREAKS; i++) breakCols.push(`Break option ${i+1}`)

  // Colonnes dynamiques pour franchise_periodes (max 3 périodes × 6 champs)
  const MAX_FRANCHISE = 3
  const franchiseCols = []
  for (let i = 0; i < MAX_FRANCHISE; i++) {
    franchiseCols.push(
      `Franchise P${i+1} - Date début`,
      `Franchise P${i+1} - Date fin`,
      `Franchise P${i+1} - Durée`,
      `Franchise P${i+1} - Montant`,
      `Franchise P${i+1} - Indexation incluse`,
      `Franchise P${i+1} - Condition`
    )
  }

  // Colonnes dynamiques pour indemnites (max 4 × 4 champs)
  const MAX_INDEM = 4
  const indemnCols = []
  for (let i = 0; i < MAX_INDEM; i++) {
    indemnCols.push(
      `Indemnité ${i+1} - Motif`,
      `Indemnité ${i+1} - Due par`,
      `Indemnité ${i+1} - Montant`,
      `Indemnité ${i+1} - Date/Condition`
    )
  }

  // Construction de la ligne
  const headers = [
    ...simpleFields.map(f => f.label),
    ...breakCols,
    ...franchiseCols,
    ...indemnCols
  ]

  const breaks = Array.isArray(data.break_options) ? data.break_options : []
  const franchise = Array.isArray(data.franchise_periodes) ? data.franchise_periodes : []
  const indem = Array.isArray(data.indemnites) ? data.indemnites : []

  const values = [
    ...simpleFields.map(f => data[f.key] ?? ''),
    ...Array.from({ length: MAX_BREAKS }, (_, i) => breaks[i] ?? ''),
    ...Array.from({ length: MAX_FRANCHISE }, (_, i) => [
      franchise[i]?.date_debut ?? '',
      franchise[i]?.date_fin ?? '',
      franchise[i]?.duree ?? '',
      franchise[i]?.montant ?? '',
      franchise[i]?.indexation_incluse ?? '',
      franchise[i]?.condition ?? '',
    ]).flat(),
    ...Array.from({ length: MAX_INDEM }, (_, i) => [
      indem[i]?.motif ?? '',
      indem[i]?.due_par ?? '',
      indem[i]?.montant ?? '',
      indem[i]?.date_limite ?? '',
    ]).flat()
  ]

  const ws = XLSX.utils.aoa_to_sheet([headers, values])
  ws['!cols'] = headers.map(() => ({ wch: 24 }))
  XLSX.utils.book_append_sheet(wb, ws, 'Base de données')
  XLSX.writeFile(wb, `lease_abstract_${(fileName||'bail').replace(/\.[^.]+$/, '')}.xlsx`)
}

function ensureArray(val) {
  if (!val) return null
  if (Array.isArray(val)) return val
  if (typeof val === 'string') {
    try { const p = JSON.parse(val); return Array.isArray(p) ? p : [val] } catch(_) { return [val] }
  }
  return null
}

function sanitizeExtracted(data) {
  if (!data || typeof data !== 'object') return data
  const d = { ...data }
  d.break_options       = ensureArray(d.break_options)
  d.surfaces_detail     = ensureArray(d.surfaces_detail)
  d.franchise_periodes  = ensureArray(d.franchise_periodes)
  d.indemnites          = ensureArray(d.indemnites)
  if (d.champs_modifies) {
    d.champs_modifies = { ...d.champs_modifies }
    d.champs_modifies.break_options      = ensureArray(d.champs_modifies.break_options)
    d.champs_modifies.surfaces_detail    = ensureArray(d.champs_modifies.surfaces_detail)
    d.champs_modifies.franchise_periodes = ensureArray(d.champs_modifies.franchise_periodes)
    d.champs_modifies.indemnites         = ensureArray(d.champs_modifies.indemnites)
  }
  return d
}

async function callClaude(base64, mediaType, prompt) {
  const res = await fetch('https://vmtmwsbebzkwxfkdpqky.supabase.co/functions/v1/hyper-action', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: 'claude-sonnet-4-6', max_tokens: 4096,
      messages: [{ role: 'user', content: [
        { type: 'document', source: { type: 'base64', media_type: mediaType, data: base64 } },
        { type: 'text', text: prompt }
      ]}]
    })
  })
  if (!res.ok) throw new Error('Erreur API: ' + res.status)
  const data = await res.json()
  // Gestion erreur API Anthropic
  if (data.type === 'error') {
    const msg = data.error?.message || 'Erreur Anthropic inconnue'
    if (msg.includes('100 PDF pages')) {
      throw new Error('Ce PDF dépasse la limite de 100 pages acceptée par Claude. Compressez-le ou découpez-le en plusieurs fichiers.')
    }
    throw new Error('Erreur Claude API : ' + msg)
  }
  let raw = ''
  if (data.content && Array.isArray(data.content)) {
    raw = data.content.map(b => (b && b.text) ? b.text : '').join('')
  } else if (data.content && typeof data.content === 'string') {
    raw = data.content
  } else if (typeof data === 'string') {
    raw = data
  } else if (data.text) {
    raw = data.text
  } else {
    throw new Error('Structure réponse inattendue : ' + JSON.stringify(data).slice(0, 300))
  }
  raw = raw.trim().replace(/```json|```/g,'').trim()
  const s = raw.indexOf('{'), e = raw.lastIndexOf('}')
  if (s === -1) throw new Error('Pas de JSON dans la réponse Claude. Réponse reçue : ' + raw.slice(0, 200))
  let jsonStr = raw.slice(s, e+1)

  function cleanJson(str) {
    // 1. Remplacer caractères typographiques
    let r = str
      .replace(/‘|’||/g, "'")
      .replace(/“|”||/g, "'")
      .replace(/–|—/g, '-')
      .replace(/ /g, ' ')
    // 2. Retirer sauts de ligne à l'intérieur des strings JSON
    let out = '', inStr = false, esc = false
    for (let i = 0; i < r.length; i++) {
      const c = r[i]
      if (esc) { out += c; esc = false; continue }
      if (c === '\\') { out += c; esc = true; continue }
      if (c === '"') { inStr = !inStr; out += c; continue }
      if (inStr && (c === '
' || c === '
' || c === '	')) { out += ' '; continue }
      out += c
    }
    return out
  }

  try { return sanitizeExtracted(JSON.parse(jsonStr)) } catch(_) {}
  try { return sanitizeExtracted(JSON.parse(cleanJson(jsonStr))) } catch(e2) {
    const pos = parseInt(e2.message.match(/position (\d+)/)?.[1] || '0')
    const ctx = cleanJson(jsonStr).slice(Math.max(0, pos - 150), pos + 150)
    throw new Error('JSON invalide pos ' + pos + ' — contexte : ' + ctx)
  }
}

// ─── Sub-components ───────────────────────────────────────────────────────────

function DropZone({ onFiles, disabled }) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef()
  const handle = useCallback(files => {
    const valid = Array.from(files).filter(f => ['pdf','docx'].includes(f.name.split('.').pop().toLowerCase()))
    if (valid.length) onFiles(valid)
    else alert('Format non supporté.')
  }, [onFiles])
  return (
    <div className={`drop-zone${dragging?' dragging':''}${disabled?' disabled':''}`}
      onClick={() => !disabled && inputRef.current?.click()}
      onDragOver={e => { e.preventDefault(); setDragging(true) }}
      onDragLeave={() => setDragging(false)}
      onDrop={e => { e.preventDefault(); setDragging(false); handle(e.dataTransfer.files) }}>
      <input ref={inputRef} type="file" accept=".pdf,.docx" multiple style={{display:'none'}} onChange={e => handle(e.target.files)} />
      <div className="drop-icon">
        <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.2">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
          <polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>
        </svg>
      </div>
      <div className="drop-title">Déposez un ou plusieurs fichiers ici</div>
      <div className="drop-sub">PDF ou DOCX · baux et avenants acceptés</div>
    </div>
  )
}

function PageLimitWarning() {
  return (
    <div style={{marginTop:'12px',padding:'10px 14px',borderRadius:'var(--r)',background:'var(--amber-bg)',border:'0.5px solid #E8C97A',fontSize:'12px',color:'var(--amber)',lineHeight:'1.6',display:'flex',gap:'8px',alignItems:'flex-start'}}>
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{flexShrink:0,marginTop:'1px'}}>
        <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
        <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
      </svg>
      <span><strong>Limite : 100 pages maximum par fichier.</strong> Si votre bail dépasse cette limite, retirez les annexes (plans, états des lieux, catalogue de charges) avant de déposer. Seules les pages de clauses sont nécessaires à l'extraction.</span>
    </div>
  )
}

function Field({ label, value, mono, verbose }) {
  return (
    <div className="field">
      <div className="field-lbl">{label}</div>
      <div className={`field-val${!value?' empty':mono?' mono':verbose?' verbose':''}`}>{value||'Non renseigné'}</div>
    </div>
  )
}

function PairBlock({ keyLabel, keyValue, keyMono, verboseLabel, verboseValue }) {
  return (
    <div className="pair-block full">
      <div className="pair-key">
        <div className="field-lbl">{keyLabel}</div>
        <div className={`field-val${!keyValue?' empty':keyMono?' mono':''}`}>{keyValue||'Non renseigné'}</div>
      </div>
      <div className="pair-verbose">
        <div className="field-lbl">{verboseLabel}</div>
        <div className={`field-val${!verboseValue?' empty':' verbose'}`}>{verboseValue||'Non renseigné'}</div>
      </div>
    </div>
  )
}

function SurfaceTable({ surfaces }) {
  const safe = Array.isArray(surfaces) ? surfaces : []
  if (!safe.length) return null
  const total = safe.reduce((acc, r) => acc + (parseFloat(String(r.surface_m2).replace(/[^0-9.]/g,'')) || 0), 0)
  return (
    <div className="field full" style={{padding:0,overflow:'hidden'}}>
      <table className="indemnites-table">
        <thead>
          <tr>
            <th>Typologie</th><th>Niveau / Localisation</th>
            <th style={{textAlign:'right'}}>Surface (m²)</th>
            <th style={{textAlign:'right'}}>Prix (€/m²)</th>
            <th style={{textAlign:'right'}}>Loyer annuel (€)</th>
          </tr>
        </thead>
        <tbody>
          {safe.map((row, i) => (
            <tr key={i}>
              <td style={{fontWeight:500}}>{row.typologie||'—'}</td>
              <td style={{color:'var(--text2)'}}>{row.localisation||'—'}</td>
              <td style={{textAlign:'right',fontWeight:500}}>{row.surface_m2||'—'}</td>
              <td style={{textAlign:'right'}}>{row.prix_unitaire||'—'}</td>
              <td style={{textAlign:'right',fontWeight:500}}>{row.loyer_annuel||'—'}</td>
            </tr>
          ))}
        </tbody>
        {total > 0 && (
          <tfoot>
            <tr style={{borderTop:'1px solid var(--border2)'}}>
              <td colSpan={2} style={{fontWeight:600,padding:'8px 10px'}}>Total</td>
              <td style={{textAlign:'right',fontWeight:600,padding:'8px 10px'}}>{total.toLocaleString('fr-FR')} m²</td>
              <td></td><td></td>
            </tr>
          </tfoot>
        )}
      </table>
    </div>
  )
}

function FranchiseTable({ periodes }) {
  const safe = Array.isArray(periodes) ? periodes : []
  if (!safe.length) return null
  return (
    <div className="field full" style={{padding:0,overflow:'hidden'}}>
      <table className="indemnites-table">
        <thead>
          <tr>
            <th>Date début</th><th>Date fin</th><th>Durée</th>
            <th>Montant exonéré</th><th>Indexation incluse</th><th>Condition</th>
          </tr>
        </thead>
        <tbody>
          {safe.map((row, i) => (
            <tr key={i}>
              <td>{row.date_debut||'—'}</td>
              <td>{row.date_fin||'—'}</td>
              <td style={{fontWeight:500}}>{row.duree||'—'}</td>
              <td style={{fontWeight:500}}>{row.montant||'—'}</td>
              <td>
                {row.indexation_incluse && (
                  <span className={`due-par ${row.indexation_incluse==='Oui'?'due-bailleur':row.indexation_incluse==='Non'?'due-preneur':''}`}>
                    {row.indexation_incluse}
                  </span>
                )}
              </td>
              <td style={{color:'var(--text2)',fontStyle:row.condition?'normal':'italic'}}>{row.condition||'—'}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function IndemniteTable({ indemnites }) {
  const safe = Array.isArray(indemnites) ? indemnites : []
  if (!safe.length) return null
  const indemnites2 = safe
  return (
    <div className="field full" style={{padding:0,overflow:'hidden'}}>
      <table className="indemnites-table">
        <thead>
          <tr>
            <th>Motif</th>
            <th>Due par</th>
            <th>Montant</th>
            <th>Date / Condition</th>
          </tr>
        </thead>
        <tbody>
          {indemnites2.map((row, i) => (
            <tr key={i}>
              <td>{row.motif||'—'}</td>
              <td>
                {row.due_par && (
                  <span className={`due-par ${row.due_par.toLowerCase().includes('preneur') ? 'due-preneur' : 'due-bailleur'}`}>
                    {row.due_par}
                  </span>
                )}
              </td>
              <td style={{fontWeight:500}}>{row.montant||'—'}</td>
              <td style={{color:'var(--text2)'}}>{row.date_limite||'—'}</td>
            </tr>
          ))}
        </tbody>
      </table>
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
          {suggestion ? `Bail détecté automatiquement (score ${Math.round(suggestion.score*100)}%) — confirmez ou choisissez.` : "Sélectionnez le bail d'origine."}
        </div>
        <div className="modal-list">
          {bails.map(b => (
            <button key={b.id} className={`modal-bail${selectedId===b.id?' sel':''}`} onClick={() => setSelectedId(b.id)}>
              <div className="modal-bail-name">{b.data?.immeuble||b.data?.adresse||b.file_name}</div>
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

function ResultsView({ item }) {
  const isAv = item.document_type === 'avenant'
  let d = isAv ? (item.data?.champs_modifies || {}) : (item.data || {})
  d = { ...d }
  if (!Array.isArray(d.break_options)) d.break_options = d.break_options ? [String(d.break_options)] : []
  if (!Array.isArray(d.franchise_periodes)) d.franchise_periodes = []
  if (!Array.isArray(d.indemnites)) d.indemnites = d.indemnites ? [d.indemnites] : []
  const meta = item.data || {}

  const breaks = Array.isArray(d.break_options) ? d.break_options : d.break_options ? [d.break_options] : []
  const indemnites = Array.isArray(d.indemnites) ? d.indemnites : null

  const pills = []
  if (d.indexation) {
    const idx = d.indexation.toLowerCase()
    if (idx.includes('ilat')) pills.push({ label:'ILAT', cls:'pill-blue' })
    else if (idx.includes('ilc')) pills.push({ label:'ILC', cls:'pill-blue' })
    else if (idx.includes('icc')) pills.push({ label:'ICC', cls:'pill-blue' })
  }
  if (d.franchise_duree) pills.push({ label: d.franchise_duree, cls:'pill-green' })

  const dateFields = [
    { key:'date_effet', label:"Prise d'effet" },
    { key:'date_signature', label:'Signature' },
    { key:'date_conge', label:'Limite congé' },
    { key:'date_fin', label:'Expiration' },
  ].filter(f => d[f.key])

  const totalDates = dateFields.length + breaks.length
  const dateCols = Math.min(Math.max(totalDates, 2), 4)

  const show = key => !isAv || d[key] != null

  return (
    <div className="result-body">
      {isAv && (
        <div className="av-banner">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{flexShrink:0,marginTop:1}}>
            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
          </svg>
          <div><strong>Avenant</strong>{meta.objet_avenant && <span style={{fontWeight:400}}> — {meta.objet_avenant}</span>}</div>
          {meta.date_effet_avenant && <span className="av-banner-date">Effet : {meta.date_effet_avenant}</span>}
        </div>
      )}

      {/* Parties */}
      {(show('preneur') || show('bailleur')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-dot"/><div className="sec-label">Parties</div></div>
          <div className="g2">
            {show('preneur') && <div className="party-card"><div className="party-role">Preneur</div><div className="party-name">{d.preneur||<span style={{color:'var(--text3)',fontStyle:'italic',fontWeight:400}}>Non renseigné</span>}</div></div>}
            {show('bailleur') && <div className="party-card"><div className="party-role">Bailleur</div><div className="party-name">{d.bailleur||<span style={{color:'var(--text3)',fontStyle:'italic',fontWeight:400}}>Non renseigné</span>}</div></div>}
            {show('garant') && d.garant && <div className="party-card full"><div className="party-role">Garant / Caution</div><div className="party-name">{d.garant}</div></div>}
          </div>
        </div>
      )}

      {/* Contrat */}
      {(show('type_bail') || show('duree_totale')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-dot"/><div className="sec-label">Contrat et durée</div></div>
          <div className="g3">
            <Field label="Type de contrat" value={d.type_bail} />
            <Field label="Durée totale" value={d.duree_totale} />
            <Field label="Durée ferme" value={d.duree_ferme} />
          </div>
        </div>
      )}

      {/* Dates */}
      {(dateFields.length > 0 || breaks.length > 0 || show('notice')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-dot"/><div className="sec-label">Dates clés</div></div>
          {(dateFields.length > 0 || breaks.length > 0) && (
            <div className="date-strip" style={{gridTemplateColumns:`repeat(${dateCols},1fr)`, marginBottom:'8px'}}>
              {dateFields.map(f => (
                <div key={f.key} className="date-card">
                  <div className="date-lbl">{f.label}</div>
                  <div className="date-val">{d[f.key]}</div>
                </div>
              ))}
              {breaks.map((br, i) => (
                <div key={i} className="date-card">
                  <div className="date-lbl"><span className="break-tag">B{breaks.length>1?i+1:''}</span> Break option</div>
                  <div className="date-val break">{br}</div>
                </div>
              ))}
            </div>
          )}
          <div className="g3">
            {show('notice') && <Field label="Préavis" value={d.notice} />}
            {show('date_limite_travaux') && d.date_limite_travaux && <Field label="Date limite travaux preneur" value={d.date_limite_travaux} />}
            {show('conditions_break') && d.conditions_break && <Field label="Conditions financières du break" value={d.conditions_break} verbose />}
          </div>
        </div>
      )}

      {/* Surfaces */}
      {(show('surface_bureaux') || show('surface_totale_m2')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-dot"/><div className="sec-label">Surfaces</div></div>
          <div className="g3">
            {show('surface_totale_m2') && (
              <div className="field">
                <div className="field-lbl">Surface totale</div>
                <div className="field-val mono">{d.surface_totale_m2 ? `${d.surface_totale_m2} m²` : '—'}</div>
              </div>
            )}
            {show('parking') && <Field label="Parking" value={d.parking} />}
            {show('rie') && d.rie && <Field label="RIE" value={d.rie} />}
          </div>
          {show('surfaces_detail') && d.surfaces_detail?.length > 0 && (
            <div style={{marginTop:'8px'}}>
              <div className="field-lbl" style={{marginBottom:'6px'}}>Surfaces louées par typologie</div>
              <SurfaceTable surfaces={d.surfaces_detail} />
            </div>
          )}
          {show('surface_bureaux') && d.surface_bureaux && (
            <div style={{marginTop:'8px'}}>
              <Field label="Détail surfaces bureaux" value={d.surface_bureaux} verbose />
            </div>
          )}
          {show('autres_surfaces') && d.autres_surfaces && (
            <div style={{marginTop:'8px'}}>
              <Field label="Autres surfaces" value={d.autres_surfaces} verbose />
            </div>
          )}
        </div>
      )}

      {/* Loyer */}
      {show('loyer_signature_montant') && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-dot"/><div className="sec-label">Loyer, taxes et charges</div></div>
          {d.loyer_signature_montant && (
            <div className="loyer-hero">
              <div>
                <div className="loyer-lbl">Loyer HT/HC annuel à la signature</div>
                <div className="loyer-amount">{fmtEur(d.loyer_signature_montant) || d.loyer_signature_montant}</div>
              </div>
              {pills.length > 0 && <div className="pills">{pills.map((p,i) => <span key={i} className={`pill ${p.cls}`}>{p.label}</span>)}</div>}
            </div>
          )}
          <div className="g2" style={{marginBottom:'8px'}}>
            {show('loyer_cours') && d.loyer_cours && <Field label="Loyer en cours" value={fmtEur(d.loyer_cours) || d.loyer_cours} />}
            {show('indexation') && <Field label="Indexation / indice" value={d.indexation} verbose />}
          </div>
          {show('loyer_signature') && d.loyer_signature && (
            <div style={{marginBottom:'8px'}}>
              <Field label="Loyer à la signature — détail complet" value={d.loyer_signature} verbose />
            </div>
          )}
          {/* Franchise tableau */}
          {show('franchise_periodes') && (d.franchise_periodes || d.franchise) && (
            <div style={{marginTop:'8px'}}>
              {Array.isArray(d.franchise_periodes) && d.franchise_periodes.length > 0 && (
                <div style={{marginBottom:'4px'}}>
                  <div className="field-lbl" style={{marginBottom:'6px'}}>Franchise — périodes</div>
                  <div className="g2"><FranchiseTable periodes={d.franchise_periodes} /></div>
                </div>
              )}
              {d.franchise && (
                <Field label="Franchise — modalités complètes" value={d.franchise} verbose />
              )}
            </div>
          )}
          <div className="g2" style={{marginTop:'8px'}}>
            {show('charges') && <Field label="Charges / TEOM" value={d.charges} verbose />}
          </div>
        </div>
      )}

      {/* Dépôt de garantie */}
      {show('depot_garantie_montant') && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-dot"/><div className="sec-label">Dépôt de garantie</div></div>
          <PairBlock
            keyLabel="Montant"
            keyValue={fmtEur(d.depot_garantie_montant) || d.depot_garantie_montant}
            keyMono
            verboseLabel="Modalités complètes"
            verboseValue={d.depot_garantie}
          />
        </div>
      )}

      {/* Participation travaux */}
      {show('travaux_montant') && (d.travaux_montant || d.travaux_modalites) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-dot"/><div className="sec-label">Participation travaux bailleur</div></div>
          <div className="g3" style={{marginBottom:'8px'}}>
            <div className="field">
              <div className="field-lbl">Montant</div>
              <div className={`field-val${!d.travaux_montant?' empty':' mono'}`}>
                {d.travaux_montant ? (fmtEur(d.travaux_montant) || d.travaux_montant) : 'Non renseigné'}
              </div>
            </div>
            {show('travaux_date_factures') && <Field label="Date limite réception factures" value={d.travaux_date_factures} />}
          </div>
          {show('travaux_modalites') && d.travaux_modalites && (
            <Field label="Modalités complètes (conditions d'appel, justificatifs, non-consommation…)" value={d.travaux_modalites} verbose />
          )}
        </div>
      )}

      {/* Indemnités */}
      {(indemnites || show('indemnites_detail')) && (d.indemnites || d.indemnites_detail) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-dot"/><div className="sec-label">Indemnités contractuelles</div></div>
          {indemnites && <div className="g2"><IndemniteTable indemnites={indemnites} /></div>}
          {show('indemnites_detail') && d.indemnites_detail && (
            <div style={{marginTop: indemnites ? '8px' : 0}}>
              <Field label="Détail verbeux" value={d.indemnites_detail} verbose />
            </div>
          )}
        </div>
      )}

      {/* Jouissance */}
      {(show('destination') || show('article_606')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-dot"/><div className="sec-label">Refacturation et jouissance</div></div>
          <div className="g2">
            {show('destination') && <Field label="Destination" value={d.destination} />}
            {show('article_606') && <Field label="Article 606" value={d.article_606} />}
            {show('sous_location') && <Field label="Sous-location" value={d.sous_location} verbose />}
            {show('cession') && <Field label="Cession" value={d.cession} verbose />}
            {show('remise_en_etat') && d.remise_en_etat && <Field label="Remise en état" value={d.remise_en_etat} verbose />}
            {show('maintenance') && d.maintenance && <Field label="Maintenance & travaux" value={d.maintenance} verbose />}
            {show('conformite') && d.conformite && <Field label="Conformité" value={d.conformite} verbose />}
            {show('accession') && d.accession && <Field label="Accession" value={d.accession} verbose />}
          </div>
        </div>
      )}
    </div>
  )
}

function HistoryPanel({ tree, onSelect, activeId, onDelete, onClear }) {
  return (
    <>
      <div className="history-list">
        {!tree.length ? <div className="history-empty">Aucune extraction sauvegardée</div>
          : tree.map(bail => (
            <div key={bail.id}>
              <div className="history-row">
                <button className={`history-btn${bail.id===activeId?' active':''}`} onClick={() => onSelect(bail)}>
                  <div className="history-name">{bail.data?.immeuble||bail.data?.adresse||bail.file_name}</div>
                  <div className="history-meta">{bail.data?.preneur?.split(',')[0]||'—'} · {formatDate(bail.created_at)}</div>
                </button>
                <button className="history-del" onClick={e => onDelete(bail,e)} title="Supprimer">✕</button>
              </div>
              {bail.avenants?.map(av => (
                <div key={av.id} className="history-row av-row">
                  <button className={`history-btn${av.id===activeId?' active':''}`} onClick={() => onSelect(av)}>
                    <div className="history-name"><span className="av-tag">A</span>{av.data?.objet_avenant||av.file_name}</div>
                    <div className="history-meta">{formatDate(av.created_at)}</div>
                  </button>
                  <button className="history-del" onClick={e => onDelete(av,e)} title="Supprimer">✕</button>
                </div>
              ))}
            </div>
          ))
        }
      </div>
      {tree.length > 0 && (
        <div className="sidebar-footer">
          <button className="btn-clear" onClick={onClear}>Vider l'historique</button>
        </div>
      )}
    </>
  )
}

// ─── App ─────────────────────────────────────────────────────────────────────

export default function App() {
  const [files,        setFiles]        = useState([])
  const [statuses,     setStatuses]     = useState([])
  const [loading,      setLoading]      = useState(false)
  const [activeItem,   setActiveItem]   = useState(null)
  const [history,      setHistory]      = useState([])
  const [histLoaded,   setHistLoaded]   = useState(false)
  const [tab,          setTab]          = useState('extract')
  const [avenantModal, setAvenantModal] = useState(null)
  const [docTypes,     setDocTypes]     = useState([]) // 'bail' | 'avenant' per file
  const [lastError,    setLastError]    = useState('')

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
  function setStatus(i, state, error) { setStatuses(prev => { const n=[...prev]; n[i]={state,error}; return n }) }

  async function extractOne(file, index, docType) {
    setStatus(index, 'loading')
    const base64 = await toBase64(file)
    const mediaType = getMediaType(file)
    const extracted = await callClaude(base64, mediaType, docType === 'avenant' ? AVENANT_PROMPT : EXTRACTION_PROMPT)
    return { extracted, docType }
  }

  async function saveExtraction(file, extracted, docType, parentId) {
    const { data: saved } = await supabase.from('extractions')
      .insert({ file_name: file.name, data: extracted, document_type: docType, parent_id: parentId || null })
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
        const docType = docTypes[i] || 'bail'
        const { extracted, docType: dt } = await extractOne(files[i], i, docType)
        if (dt === 'avenant') {
          const match = findBestMatch(extracted?.bail_reference, allBails)
          await new Promise(resolve => setAvenantModal({ index: i, file: files[i], extracted, suggestion: match, resolve }))
        } else {
          const saved = await saveExtraction(files[i], extracted, 'bail', null)
          if (saved) { setActiveItem(saved); setHistory(prev => [{ ...saved, avenants: [] }, ...prev]) }
          setStatus(i, 'done')
        }
      } catch(e) { setStatus(i, 'error', e.message); setLastError(e.message) }
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

  async function handleDeleteItem(item, e) {
    e.stopPropagation()
    const label = item.data?.immeuble || item.data?.adresse || item.file_name
    if (!window.confirm(`Supprimer "${label}" ? Cette action est irréversible.`)) return
    await supabase.from('extractions').delete().eq('id', item.id)
    if (activeItem?.id === item.id) setActiveItem(null)
    setHistory(prev => prev.filter(b => b.id !== item.id).map(b => ({ ...b, avenants: (b.avenants||[]).filter(a => a.id !== item.id) })))
  }

  async function handleClearHistory() {
    if (!window.confirm("Vider tout l'historique ? Cette action est irréversible.")) return
    await supabase.from('extractions').delete().neq('id', '00000000-0000-0000-0000-000000000000')
    setHistory([]); setHistLoaded(false); setActiveItem(null)
  }

  function handleClear() { setFiles([]); setStatuses([]); setActiveItem(null); setDocTypes([]); setLastError('') }

  const d = activeItem?.data || {}
  const resultTitle = d.immeuble || d.adresse || activeItem?.file_name || ''
  const resultSub = [d.preneur?.split(',')[0], d.bailleur?.split(',')[0], d.date_signature ? `Signé le ${d.date_signature}` : null].filter(Boolean).join(' · ')

  return (
    <>
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
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14 2 14 8 20 8"/>
              <line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>
            </svg>
            Lease Reader
          </div>
          <nav className="sidebar-nav">
            <button className={`nav-item${tab==='extract'?' active':''}`} onClick={() => switchTab('extract')}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
              </svg>
              Extraire
            </button>
            <button className={`nav-item${tab==='history'?' active':''}`} onClick={() => switchTab('history')}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
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
              onDelete={handleDeleteItem}
              onClear={handleClearHistory}
            />
          )}
        </aside>

        <main className="main">
          {activeItem && (
            <div className="result-topbar">
              <div className="result-tag">
                <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>
                </svg>
                {activeItem.document_type === 'avenant' ? 'Avenant' : 'Bail commercial'}
              </div>
              <div className="result-title">{resultTitle}</div>
              {resultSub && <div className="result-sub">{resultSub}</div>}
              <div className="result-actions">
                <button className="btn back" onClick={handleClear}>← Nouvelle extraction</button>
                <button className="btn primary" onClick={() => exportToExcel(
                  activeItem.document_type==='avenant' ? activeItem.data?.champs_modifies||{} : activeItem.data||{},
                  activeItem.file_name
                )}>
                  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                    <polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                  </svg>
                  Exporter Excel
                </button>
              </div>
            </div>
          )}

          <div className="content">
            {!activeItem && (
              <div className="extract-wrap">
                <DropZone onFiles={setFiles} disabled={loading} />
                <PageLimitWarning />
                <div className="file-queue">
                  {files.map((f, i) => {
                    const st = statuses[i] || {}
                    return (
                      <div key={i} className={`queue-item ${st.state||''}`}>
                        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
                        </svg>
                        <span className="queue-name">{f.name}</span>
                        <span className="queue-size">({(f.size/1024).toFixed(0)} Ko)</span>
                        {st.state==='loading' && <span className="queue-status">En cours…</span>}
                        {st.state==='done'    && <span className="queue-status ok">✓ Extrait</span>}
                        {st.state==='error'   && <span className="queue-status err" title={st.error}>✕ Erreur</span>}
                        {!st.state && (
                        <>
                          <div style={{display:'flex',gap:0,border:'0.5px solid var(--border2)',borderRadius:'6px',overflow:'hidden',flexShrink:0}}>
                            <button
                              style={{padding:'3px 8px',fontSize:'11px',fontWeight:500,border:'none',background:(!docTypes[i]||docTypes[i]==='bail')?'var(--accent)':'transparent',color:(!docTypes[i]||docTypes[i]==='bail')?'#fff':'var(--text2)',cursor:'pointer'}}
                              onClick={() => setDocTypes(p => { const n=[...p]; n[i]='bail'; return n })}>Bail</button>
                            <button
                              style={{padding:'3px 8px',fontSize:'11px',fontWeight:500,border:'none',borderLeft:'0.5px solid var(--border2)',background:docTypes[i]==='avenant'?'var(--accent)':'transparent',color:docTypes[i]==='avenant'?'#fff':'var(--text2)',cursor:'pointer'}}
                              onClick={() => setDocTypes(p => { const n=[...p]; n[i]='avenant'; return n })}>Avenant</button>
                          </div>
                          <button className="queue-remove" onClick={() => { setFiles(p=>p.filter((_,j)=>j!==i)); setStatuses(p=>p.filter((_,j)=>j!==i)); setDocTypes(p=>p.filter((_,j)=>j!==i)) }}>✕</button>
                        </>
                      )}
                      </div>
                    )
                  })}
                </div>
                {files.length > 0 && !loading && (
                  <div className="extract-bar">
                    <button className="btn primary" onClick={handleExtract}>
                      Extraire {files.length > 1 ? `les ${files.length} fichiers` : 'le fichier'}
                    </button>
                    <button className="btn" onClick={handleClear}>Tout effacer</button>
                  </div>
                )}
                {lastError && (
                  <div style={{marginTop:'12px',padding:'12px 14px',borderRadius:'var(--r)',background:'var(--danger-bg)',border:'0.5px solid #F09595',fontSize:'12px',color:'var(--danger)',lineHeight:'1.6'}}>
                    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',gap:'8px'}}>
                      <strong>Erreur :</strong>
                      <button onClick={() => {navigator.clipboard.writeText(lastError)}} style={{background:'none',border:'0.5px solid var(--danger)',borderRadius:'4px',color:'var(--danger)',fontSize:'11px',padding:'2px 6px',cursor:'pointer',flexShrink:0}}>Copier</button>
                    </div>
                    <div style={{marginTop:'4px',wordBreak:'break-all'}}>{lastError}</div>
                  </div>
                )}
                {loading && (
                  <div>
                    <div className="progress-track"><div className="progress-bar active"/></div>
                    <div className="status-msg">Extraction en cours…</div>
                  </div>
                )}
              </div>
            )}
            {activeItem && <ResultsView item={activeItem} />}
          </div>
        </main>
      </div>
    </>
  )
}
