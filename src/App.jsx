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
    { key: 'surface_totale_m2', label: 'Surface totale (m²)' },
    { key: 'surfaces_detail', label: 'Tableau surfaces par typologie' },
    { key: 'parking_nb_places', label: 'Parking — nombre de places' },
    { key: 'rie', label: 'RIE' },
  ]},
  { id: 'loyer', label: 'Loyer', fields: [
    { key: 'loyer_signature_montant', label: 'Loyer signature (€/an)' },
    { key: 'loyer_signature', label: 'Loyer signature — détail' },
    { key: 'loyer_cours', label: 'Loyer de base' },
    { key: 'indexation', label: 'Indexation / indice' },
    { key: 'franchise_periodes', label: 'Franchise — périodes' },
    { key: 'franchise', label: 'Franchise — modalités' },
    { key: 'charges', label: 'Charges / TEOM' },
  ]},
  { id: 'depot', label: 'Dépôt de garantie', fields: [
    { key: 'depot_garantie_montant', label: 'Dépôt de garantie (€)' },
    { key: 'depot_garantie', label: 'Dépôt de garantie — modalités' },
  ]},
  { id: 'travaux', label: 'Participation travaux bailleur', fields: [
    { key: 'travaux_montant', label: 'Montant (€)' },
    { key: 'travaux_date_factures', label: 'Date limite réception factures' },
    { key: 'travaux_modalites', label: 'Modalités complètes' },
  ]},
  { id: 'indemnites', label: 'Indemnités contractuelles', fields: [
    { key: 'indemnites', label: 'Tableau des indemnités' },
    { key: 'indemnites_detail', label: 'Détail' },
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

const EXTRACTION_PROMPT = `Expert baux commerciaux français. Extrais les données du document. Retourne UNIQUEMENT du JSON minifié sur UNE SEULE LIGNE, sans indentation, sans saut de ligne, sans markdown.

REGLES: Guillemets droits ASCII uniquement. Pas de retour a la ligne dans les valeurs. Echappe les guillemets internes avec backslash. Champs _montant=chiffres bruts sans symbole ni espace (ex: 123405.50). null si absent. DISTINCTION IMPORTANTE: duree_totale=duree totale du bail (date_effet -> date_fin). duree_ferme=calcule comme suit: si break_options existe, duree_ferme = intervalle entre date_effet et la PREMIERE break option (ex: effet 01/09/2025, premiere break 31/08/2028 -> duree_ferme="3 ans"). Si pas de break_options, duree_ferme = duree_totale. Si le bail mentionne explicitement une duree ferme, utilise cette valeur.

{"adresse":null,"immeuble":null,"ville":null,"type_bail":null,"duree_totale":null,"duree_ferme":null,"preneur":null,"bailleur":null,"garant":null,"date_effet":null,"date_signature":null,"break_options":[],"notice":null,"date_conge":null,"date_fin":null,"date_limite_travaux":null,"conditions_break":null,"surface_totale_m2":null,"surfaces_detail":[],"parking_nb_places":null,"parking":null,"rie":null,"loyer_signature_montant":null,"loyer_signature":null,"loyer_cours":null,"indexation":null,"franchise_periodes":[],"franchise":null,"charges":null,"depot_garantie_montant":null,"depot_garantie":null,"travaux_montant":null,"travaux_date_factures":null,"travaux_modalites":null,"indemnites":[],"indemnites_detail":null,"article_606":null,"conformite":null,"accession":null,"remise_en_etat":null,"maintenance":null,"destination":null,"sous_location":null,"cession":null}

Formats:
- surfaces_detail: [{"categorie":"Bureaux","niveau":"5eme etage","surface_m2":"2224.98","prix_unitaire":"290","loyer_annuel":"645244"}] — categorie=Bureaux/Archives/Stationnement/Commerce/RIE/Autres
- break_options: ["31/08/2027","31/08/2030"]
- franchise_periodes: [{"date_debut":"jj/mm/aaaa","date_fin":"jj/mm/aaaa","duree":"3 mois","montant":"123405","indexation_incluse":"Oui/Non/Non precise","condition":null}] — montant=chiffres bruts sans symbole. Pour indexation_incluse: "Non" si la franchise est calculee sur le "loyer de base" (meme indexe dans le temps, le terme "loyer de base" signifie hors indexation supplementaire); "Oui" uniquement si le texte dit explicitement "loyer indexe" ou "loyer en cours"; "Non precise" si aucune reference au calcul n'est donnee
- indemnites: UNIQUEMENT les indemnites financieres conditionnees a l'exercice ou non d'une option (break option, renouvellement, fin de bail). Exemples valides: restitution de franchise si depart a une break, indemnite liberatoire de remise en etat si depart avant terme, complement de franchise si maintien au-dela d'une echeance. EXCLURE ABSOLUMENT: honoraires, frais d'acte, cautionnements, penalites de retard, indemnites d'occupation, provisions. Si aucune indemnite ne correspond strictement a ce critere, retourne un tableau vide []. Format: [{"motif":"...","due_par":"Preneur ou Bailleur","montant":"chiffres bruts","date_limite":"..."}]
- parking_nb_places: decompte exact ex: "114 places (98 interieures + 16 exterieures)"`

const AVENANT_PROMPT = `Expert baux commerciaux français. Ce document est un AVENANT. JSON minifie sur UNE SEULE LIGNE, sans markdown.

REGLES: Ne renseigne dans champs_modifies QUE les champs modifies. null pour les autres. Champs _montant=chiffres bruts sans symbole.

{"bail_reference":{"preneur":null,"bailleur":null,"date_bail_origine":null,"adresse":null,"immeuble":null},"date_effet_avenant":null,"date_signature_avenant":null,"objet_avenant":null,"champs_modifies":{"adresse":null,"immeuble":null,"ville":null,"type_bail":null,"duree_totale":null,"duree_ferme":null,"preneur":null,"bailleur":null,"garant":null,"date_effet":null,"date_signature":null,"break_options":null,"notice":null,"date_conge":null,"date_fin":null,"date_limite_travaux":null,"conditions_break":null,"surface_totale_m2":null,"surfaces_detail":null,"parking_nb_places":null,"parking":null,"rie":null,"loyer_signature_montant":null,"loyer_signature":null,"loyer_cours":null,"indexation":null,"franchise_periodes":null,"franchise":null,"charges":null,"depot_garantie_montant":null,"depot_garantie":null,"travaux_montant":null,"travaux_date_factures":null,"travaux_modalites":null,"indemnites":null,"indemnites_detail":null,"article_606":null,"conformite":null,"accession":null,"remise_en_etat":null,"maintenance":null,"destination":null,"sous_location":null,"cession":null}}`


const DETECT_PROMPT = `Analyse ce document immobilier. Reponds UNIQUEMENT avec ce JSON sur une ligne:
{"type":"bail","pertinent":true,"raison":"","preneur":"","bailleur":"","adresse":"","immeuble":""}
Regles:
- type: "bail" si bail original, "avenant" si avenant ou avenant rectificatif
- pertinent: true si document est un bail/avenant valide et le bail semble actif ou potentiellement actif (un avenant peut prolonger un bail expire -> pertinent:true). false si ce n'est pas un bail/avenant, ou si le bail est clairement expire et sans prolongation, ou si le document est illisible/hors sujet
- raison: courte explication seulement si pertinent:false
- preneur, bailleur, adresse, immeuble: extrais ces valeurs meme partiellement, elles servent a identifier le bail associe (pour un avenant)`

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
  return file.name.toLowerCase().endsWith('.pdf')
    ? 'application/pdf'
    : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
}

function formatDate(iso) {
  if (!iso) return '—'
  return new Date(iso).toLocaleDateString('fr-FR', { day: '2-digit', month: 'short', year: 'numeric' })
}

// Parse a raw montant string to a float number (strips currency symbols, spaces)
function parseAmount(val) {
  if (!val) return null
  const n = parseFloat(String(val).replace(/[^0-9.,]/g, '').replace(',', '.'))
  return isNaN(n) ? null : n
}

// Format for display
function fmtEur(val) {
  const n = parseAmount(val)
  if (n === null) return val || null
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

// ─── Excel export ─────────────────────────────────────────────────────────────

function exportToExcel(data, fileName) {
  const wb = XLSX.utils.book_new()

  const simpleFields = ALL_FIELDS.filter(f => !['break_options','franchise_periodes','indemnites','surfaces_detail'].includes(f.key))

  const MAX_BREAKS = 4
  const MAX_FRANCHISE = 5
  const MAX_INDEM = 5

  const breakCols = Array.from({ length: MAX_BREAKS }, (_, i) => `Break option ${i+1}`)

  const franchiseCols = Array.from({ length: MAX_FRANCHISE }, (_, i) => [
    `Franchise P${i+1} - Date debut`,
    `Franchise P${i+1} - Date fin`,
    `Franchise P${i+1} - Duree`,
    `Franchise P${i+1} - Montant (EUR)`,
    `Franchise P${i+1} - Indexation incluse`,
    `Franchise P${i+1} - Condition`,
  ]).flat()

  const indemnCols = Array.from({ length: MAX_INDEM }, (_, i) => [
    `Indemnite ${i+1} - Motif`,
    `Indemnite ${i+1} - Due par`,
    `Indemnite ${i+1} - Montant (EUR)`,
    `Indemnite ${i+1} - Date/Condition`,
  ]).flat()

  const headers = [...simpleFields.map(f => f.label), ...breakCols, ...franchiseCols, ...indemnCols]

  const breaks   = Array.isArray(data.break_options)    ? data.break_options    : []
  const franchise = Array.isArray(data.franchise_periodes) ? data.franchise_periodes : []
  const indem    = Array.isArray(data.indemnites)        ? data.indemnites       : []

  const values = [
    ...simpleFields.map(f => {
      const v = data[f.key]
      // Champs montants : exporter en nombre
      if (f.key.includes('_montant') || f.key === 'travaux_montant') {
        const n = parseAmount(v)
        return n !== null ? n : (v ?? '')
      }
      return v ?? ''
    }),
    ...Array.from({ length: MAX_BREAKS }, (_, i) => breaks[i] ?? ''),
    ...Array.from({ length: MAX_FRANCHISE }, (_, i) => [
      franchise[i]?.date_debut ?? '',
      franchise[i]?.date_fin ?? '',
      franchise[i]?.duree ?? '',
      parseAmount(franchise[i]?.montant) ?? franchise[i]?.montant ?? '', // nombre
      franchise[i]?.indexation_incluse ?? '',
      franchise[i]?.condition ?? '',
    ]).flat(),
    ...Array.from({ length: MAX_INDEM }, (_, i) => [
      indem[i]?.motif ?? '',
      indem[i]?.due_par ?? '',
      parseAmount(indem[i]?.montant) ?? indem[i]?.montant ?? '', // nombre
      indem[i]?.date_limite ?? '',
    ]).flat(),
  ]

  const ws = XLSX.utils.aoa_to_sheet([headers, values])

  // Formatter les colonnes montant en format nombre Excel
  const montantColIndices = []
  headers.forEach((h, idx) => {
    if (h.includes('(EUR)') || h.includes('montant') || h.includes('Montant')) {
      montantColIndices.push(idx)
    }
  })
  montantColIndices.forEach(colIdx => {
    const cellAddr = XLSX.utils.encode_cell({ r: 1, c: colIdx })
    if (ws[cellAddr] && typeof ws[cellAddr].v === 'number') {
      ws[cellAddr].t = 'n'
      ws[cellAddr].z = '#,##0'
    }
  })

  ws['!cols'] = headers.map(() => ({ wch: 22 }))
  XLSX.utils.book_append_sheet(wb, ws, 'Base de données')
  XLSX.writeFile(wb, `lease_abstract_${(fileName || 'bail').replace(/\.[^.]+$/, '')}.xlsx`)
}

// ─── JSON cleaning & parsing ──────────────────────────────────────────────────

function ensureArray(val) {
  if (!val) return null
  if (Array.isArray(val)) return val
  if (typeof val === 'string') {
    try { const p = JSON.parse(val); return Array.isArray(p) ? p : [val] } catch (_) { return [val] }
  }
  return null
}

function sanitizeExtracted(data) {
  if (!data || typeof data !== 'object') return data
  const d = { ...data }
  d.break_options      = ensureArray(d.break_options)
  d.surfaces_detail    = ensureArray(d.surfaces_detail)
  d.franchise_periodes = ensureArray(d.franchise_periodes)
  d.indemnites         = ensureArray(d.indemnites)
  if (d.champs_modifies) {
    d.champs_modifies = { ...d.champs_modifies }
    d.champs_modifies.break_options      = ensureArray(d.champs_modifies.break_options)
    d.champs_modifies.surfaces_detail    = ensureArray(d.champs_modifies.surfaces_detail)
    d.champs_modifies.franchise_periodes = ensureArray(d.champs_modifies.franchise_periodes)
    d.champs_modifies.indemnites         = ensureArray(d.champs_modifies.indemnites)
  }
  return d
}

function cleanJson(str) {
  let out = ''
  let inStr = false
  let esc = false
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i)
    const c = str[i]
    if (esc) { out += c; esc = false; continue }
    if (inStr && code === 92) { out += c; esc = true; continue } // backslash
    if (code === 34) { // guillemet
      if (!inStr) { inStr = true; out += c; continue }
      let j = i + 1
      while (j < str.length && (str.charCodeAt(j) === 32 || str.charCodeAt(j) === 10 || str.charCodeAt(j) === 13 || str.charCodeAt(j) === 9)) j++
      const nc = str.charCodeAt(j)
      if (nc === 58 || nc === 44 || nc === 125 || nc === 93 || j >= str.length) {
        inStr = false; out += c
      } else {
        out += '\\"'
      }
      continue
    }
    if (inStr) {
      if (code === 10 || code === 13 || code === 9) { out += ' '; continue }
      if (code === 8216 || code === 8217) { out += "'"; continue }
      if (code === 8220 || code === 8221) { out += "'"; continue }
      if (code === 8211 || code === 8212) { out += '-'; continue }
      if (code === 160) { out += ' '; continue }
    }
    out += c
  }
  return out
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
  if (data.type === 'error') {
    const msg = data.error?.message || 'Erreur Anthropic'
    if (msg.includes('100 PDF pages')) throw new Error('PDF > 100 pages : retirez les annexes avant de déposer.')
    throw new Error('Claude API : ' + msg)
  }
  let raw = ''
  if (data.content && Array.isArray(data.content)) raw = data.content.map(b => b?.text || '').join('')
  else if (data.text) raw = data.text
  else if (typeof data === 'string') raw = data
  else throw new Error('Réponse inattendue : ' + JSON.stringify(data).slice(0, 200))
  raw = raw.trim().replace(/```json|```/g, '').trim()
  const s = raw.indexOf('{'), e = raw.lastIndexOf('}')
  if (s === -1) throw new Error('Pas de JSON. Reçu : ' + raw.slice(0, 200))
  const jsonStr = raw.slice(s, e + 1)
  try { return sanitizeExtracted(JSON.parse(jsonStr)) } catch (_) {}
  const cleaned = cleanJson(jsonStr)
  try { return sanitizeExtracted(JSON.parse(cleaned)) } catch (e2) {
    const pos = parseInt(e2.message.match(/position (\d+)/)?.[1] || '0')
    throw new Error('JSON pos ' + pos + ' : ' + cleaned.slice(Math.max(0, pos - 250), pos + 100))
  }
}

// ─── Sub-components ───────────────────────────────────────────────────────────

function DropZone({ onFiles, disabled }) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef()
  const handle = useCallback(files => {
    const valid = Array.from(files).filter(f => ['pdf', 'docx'].includes(f.name.split('.').pop().toLowerCase()))
    if (valid.length) onFiles(valid)
    else alert('Format non supporté. PDF ou DOCX uniquement.')
  }, [onFiles])
  return (
    <div
      className={`drop-zone${dragging ? ' dragging' : ''}${disabled ? ' disabled' : ''}`}
      onClick={() => !disabled && inputRef.current?.click()}
      onDragOver={e => { e.preventDefault(); setDragging(true) }}
      onDragLeave={() => setDragging(false)}
      onDrop={e => { e.preventDefault(); setDragging(false); handle(e.dataTransfer.files) }}
    >
      <input ref={inputRef} type="file" accept=".pdf,.docx" multiple style={{ display: 'none' }} onChange={e => handle(e.target.files)} />
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
    <div className="warning-box">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ flexShrink: 0, marginTop: '1px' }}>
        <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
        <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
      </svg>
      <span><strong>Limite : 100 pages maximum par fichier.</strong> Retirez les annexes (plans, états des lieux, catalogue de charges) si nécessaire.</span>
    </div>
  )
}

function Field({ label, value, mono, verbose }) {
  return (
    <div className="field">
      <div className="field-lbl">{label}</div>
      <div className={`field-val${!value ? ' empty' : mono ? ' mono' : verbose ? ' verbose' : ''}`}>{value || 'Non renseigné'}</div>
    </div>
  )
}

function PairBlock({ keyLabel, keyValue, keyMono, verboseLabel, verboseValue }) {
  return (
    <div className="pair-block full">
      <div className="pair-key">
        <div className="field-lbl">{keyLabel}</div>
        <div className={`field-val${!keyValue ? ' empty' : keyMono ? ' mono' : ''}`}>{keyValue || 'Non renseigné'}</div>
      </div>
      <div className="pair-verbose">
        <div className="field-lbl">{verboseLabel}</div>
        <div className={`field-val${!verboseValue ? ' empty' : ' verbose'}`}>{verboseValue || 'Non renseigné'}</div>
      </div>
    </div>
  )
}

function SurfaceTable({ surfaces }) {
  const safe = Array.isArray(surfaces) ? surfaces : []
  if (!safe.length) return null
  const total = safe.reduce((acc, r) => {
    const cat = (r.categorie || r.typologie || '').toLowerCase()
    if (cat.includes('station') || cat.includes('parking') || cat.includes('place')) return acc
    return acc + (parseFloat(String(r.surface_m2).replace(/[^0-9.]/g, '')) || 0)
  }, 0)
  return (
    <div className="table-wrap">
      <table className="indemnites-table">
        <thead>
          <tr>
            <th>Catégorie</th><th>Niveau / Localisation</th>
            <th style={{ textAlign: 'right' }}>Surface (m²)</th>
            <th style={{ textAlign: 'right' }}>Prix (€/m²)</th>
            <th style={{ textAlign: 'right' }}>Loyer annuel (€)</th>
          </tr>
        </thead>
        <tbody>
          {safe.map((row, i) => (
            <tr key={i}>
              <td style={{ fontWeight: 500 }}>{row.categorie || row.typologie || '—'}</td>
              <td style={{ color: 'var(--text2)' }}>{row.niveau || row.localisation || '—'}</td>
              <td style={{ textAlign: 'right', fontWeight: 500 }}>{row.surface_m2 || '—'}</td>
              <td style={{ textAlign: 'right' }}>{row.prix_unitaire ? fmtEur(row.prix_unitaire) : '—'}</td>
              <td style={{ textAlign: 'right', fontWeight: 500 }}>{row.loyer_annuel ? fmtEur(row.loyer_annuel) : '—'}</td>
            </tr>
          ))}
        </tbody>
        {total > 0 && (
          <tfoot>
            <tr style={{ borderTop: '1px solid var(--border2)' }}>
              <td colSpan={2} style={{ fontWeight: 600, padding: '8px 10px' }}>Total bureaux + archives</td>
              <td style={{ textAlign: 'right', fontWeight: 600, padding: '8px 10px' }}>{total.toLocaleString('fr-FR')} m²</td>
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
    <div className="table-wrap">
      <table className="indemnites-table">
        <thead>
          <tr>
            <th>Date début</th><th>Date fin</th><th>Durée</th>
            <th style={{ textAlign: 'right' }}>Montant exonéré</th>
            <th>Indexation incluse</th><th>Condition</th>
          </tr>
        </thead>
        <tbody>
          {safe.map((row, i) => (
            <tr key={i}>
              <td>{row.date_debut || '—'}</td>
              <td>{row.date_fin || '—'}</td>
              <td style={{ fontWeight: 500, whiteSpace: 'nowrap' }}>{row.duree || '—'}</td>
              <td style={{ textAlign: 'right', fontWeight: 500 }}>
                {row.montant ? fmtEur(row.montant) : '—'}
              </td>
              <td>
                {row.indexation_incluse && (
                  <span className={`due-par ${row.indexation_incluse === 'Oui' ? 'due-bailleur' : row.indexation_incluse === 'Non' ? 'due-preneur' : ''}`}>
                    {row.indexation_incluse}
                  </span>
                )}
              </td>
              <td style={{ color: 'var(--text2)', fontStyle: row.condition ? 'normal' : 'italic' }}>{row.condition || '—'}</td>
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
  return (
    <div className="table-wrap">
      <table className="indemnites-table">
        <thead>
          <tr>
            <th>Motif</th><th>Due par</th>
            <th style={{ textAlign: 'right' }}>Montant</th>
            <th>Date / Condition</th>
          </tr>
        </thead>
        <tbody>
          {safe.map((row, i) => (
            <tr key={i}>
              <td>{row.motif || '—'}</td>
              <td>
                {row.due_par && (
                  <span className={`due-par ${row.due_par.toLowerCase().includes('preneur') ? 'due-preneur' : 'due-bailleur'}`}>
                    {row.due_par}
                  </span>
                )}
              </td>
              <td style={{ textAlign: 'right', fontWeight: 500 }}>
                {row.montant ? fmtEur(row.montant) : '—'}
              </td>
              <td style={{ color: 'var(--text2)' }}>{row.date_limite || '—'}</td>
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
          {suggestion ? `Bail détecté automatiquement (score ${Math.round(suggestion.score * 100)}%) — confirmez ou choisissez.` : "Sélectionnez le bail d'origine."}
        </div>
        <div className="modal-list">
          {bails.map(b => (
            <button key={b.id} className={`modal-bail${selectedId === b.id ? ' sel' : ''}`} onClick={() => setSelectedId(b.id)}>
              <div className="modal-bail-name">{b.data?.immeuble || b.data?.adresse || b.file_name}</div>
              <div className="modal-bail-meta">{b.data?.preneur || '—'} · {formatDate(b.created_at)}</div>
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
  if (!Array.isArray(d.surfaces_detail)) d.surfaces_detail = []
  const meta = item.data || {}

  const breaks = d.break_options
  const indemnites = d.indemnites.length > 0 ? d.indemnites : null

  const pills = []
  if (d.indexation) {
    const idx = d.indexation.toLowerCase()
    if (idx.includes('ilat')) pills.push({ label: 'ILAT', cls: 'pill-blue' })
    else if (idx.includes('ilc')) pills.push({ label: 'ILC', cls: 'pill-blue' })
    else if (idx.includes('icc')) pills.push({ label: 'ICC', cls: 'pill-blue' })
  }
  if (d.franchise_periodes?.length > 0) {
    const totalMois = d.franchise_periodes.reduce((acc, p) => acc + (parseInt(p.duree) || 0), 0)
    pills.push({ label: totalMois > 0 ? `Franchise ${totalMois} mois` : 'Franchise', cls: 'pill-green' })
  } else if (d.franchise) {
    pills.push({ label: 'Franchise', cls: 'pill-green' })
  }

  const dateFields = [
    { key: 'date_effet', label: "Prise d'effet" },
    { key: 'date_signature', label: 'Signature' },
    { key: 'date_conge', label: 'Limite congé' },
    { key: 'date_fin', label: 'Expiration' },
  ].filter(f => d[f.key])

  const totalDates = dateFields.length + breaks.length
  const dateCols = Math.min(Math.max(totalDates, 2), 4)
  const show = key => !isAv || d[key] != null

  return (
    <div className="result-body">
      {isAv && (
        <div className="av-banner">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ flexShrink: 0, marginTop: 1 }}>
            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
          </svg>
          <div><strong>Avenant</strong>{meta.objet_avenant && <span style={{ fontWeight: 400 }}> — {meta.objet_avenant}</span>}</div>
          {meta.date_effet_avenant && <span className="av-banner-date">Effet : {meta.date_effet_avenant}</span>}
        </div>
      )}

      {/* Parties */}
      {(show('preneur') || show('bailleur')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Parties</div></div>
          <div className="g2">
            {show('preneur') && <div className="party-card"><div className="party-role">Preneur</div><div className="party-name">{d.preneur || <span style={{ color: 'var(--text3)', fontStyle: 'italic', fontWeight: 400 }}>Non renseigné</span>}</div></div>}
            {show('bailleur') && <div className="party-card"><div className="party-role">Bailleur</div><div className="party-name">{d.bailleur || <span style={{ color: 'var(--text3)', fontStyle: 'italic', fontWeight: 400 }}>Non renseigné</span>}</div></div>}
            {show('garant') && d.garant && <div className="party-card full"><div className="party-role">Garant / Caution</div><div className="party-name">{d.garant}</div></div>}
          </div>
        </div>
      )}

      {/* Contrat */}
      {(show('type_bail') || show('duree_totale')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Contrat et durée</div></div>
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
          <div className="sec-hd"><div className="sec-label">Dates clés</div></div>
          {(dateFields.length > 0 || breaks.length > 0) && (
            <div className="date-strip" style={{ gridTemplateColumns: `repeat(${dateCols},1fr)`, marginBottom: '8px' }}>
              {dateFields.map(f => (
                <div key={f.key} className="date-card">
                  <div className="date-lbl">{f.label}</div>
                  <div className="date-val">{d[f.key]}</div>
                </div>
              ))}
              {breaks.map((br, i) => (
                <div key={i} className="date-card">
                  <div className="date-lbl"><span className="break-tag">B{breaks.length > 1 ? i + 1 : ''}</span> Break option</div>
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
      {(show('surfaces_detail') || show('surface_totale_m2')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Surfaces</div></div>
          <div className="g3">
            {show('surface_totale_m2') && (
              <div className="field">
                <div className="field-lbl">Surface totale</div>
                <div className="field-val mono">{d.surface_totale_m2 ? `${d.surface_totale_m2} m²` : '—'}</div>
              </div>
            )}
            {show('parking_nb_places') && (
              <div className="field">
                <div className="field-lbl">Nombre de places de parking</div>
                <div className={`field-val${!d.parking_nb_places ? ' empty' : ' mono'}`}>{d.parking_nb_places || 'Non renseigné'}</div>
              </div>
            )}
            {show('rie') && d.rie && <Field label="RIE" value={d.rie} />}
          </div>
          {d.surfaces_detail?.length > 0 && (
            <div style={{ marginTop: '8px' }}>
              <div className="field-lbl" style={{ marginBottom: '6px' }}>Surfaces louées par typologie</div>
              <SurfaceTable surfaces={d.surfaces_detail} />
            </div>
          )}

        </div>
      )}

      {/* Loyer */}
      {show('loyer_signature_montant') && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Loyer, taxes et charges</div></div>
          {d.loyer_signature_montant && (
            <div className="loyer-hero">
              <div>
                <div className="loyer-lbl">Loyer HT/HC annuel à la signature</div>
                <div className="loyer-amount">{fmtEur(d.loyer_signature_montant) || d.loyer_signature_montant}</div>
              </div>
              {pills.length > 0 && <div className="pills">{pills.map((p, i) => <span key={i} className={`pill ${p.cls}`}>{p.label}</span>)}</div>}
            </div>
          )}
          <div className="g2" style={{ marginBottom: '8px' }}>
            {show('loyer_cours') && d.loyer_cours && <Field label="Loyer de base" value={fmtEur(d.loyer_cours) || d.loyer_cours} />}
            {show('indexation') && <Field label="Indexation / indice" value={d.indexation} verbose />}
          </div>
          {show('loyer_signature') && d.loyer_signature && (
            <div style={{ marginBottom: '8px' }}>
              <Field label="Loyer à la signature — détail complet" value={d.loyer_signature} verbose />
            </div>
          )}
          {(d.franchise_periodes?.length > 0 || d.franchise) && (
            <div style={{ marginTop: '8px' }}>
              {d.franchise_periodes?.length > 0 && (
                <div style={{ marginBottom: '4px' }}>
                  <div className="field-lbl" style={{ marginBottom: '6px' }}>Franchise — périodes</div>
                  <FranchiseTable periodes={d.franchise_periodes} />
                </div>
              )}
              {d.franchise && <Field label="Franchise — modalités complètes" value={d.franchise} verbose />}
            </div>
          )}
          <div className="g2" style={{ marginTop: '8px' }}>
            {show('charges') && <Field label="Charges / TEOM" value={d.charges} verbose />}
          </div>
        </div>
      )}

      {/* Dépôt de garantie */}
      {show('depot_garantie_montant') && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Dépôt de garantie</div></div>
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
          <div className="sec-hd"><div className="sec-label">Participation travaux bailleur</div></div>
          <div className="g3" style={{ marginBottom: '8px' }}>
            <div className="field">
              <div className="field-lbl">Montant</div>
              <div className={`field-val${!d.travaux_montant ? ' empty' : ' mono'}`}>
                {d.travaux_montant ? (fmtEur(d.travaux_montant) || d.travaux_montant) : 'Non renseigné'}
              </div>
            </div>
            {show('travaux_date_factures') && <Field label="Date limite réception factures" value={d.travaux_date_factures} />}
          </div>
          {show('travaux_modalites') && d.travaux_modalites && (
            <Field label="Modalités complètes" value={d.travaux_modalites} verbose />
          )}
        </div>
      )}

      {/* Indemnités */}
      {(indemnites || d.indemnites_detail) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Indemnités contractuelles liées aux échéances</div></div>
          {indemnites && <IndemniteTable indemnites={indemnites} />}
          {show('indemnites_detail') && d.indemnites_detail && (
            <div style={{ marginTop: indemnites ? '8px' : 0 }}>
              <Field label="Détail" value={d.indemnites_detail} verbose />
            </div>
          )}
        </div>
      )}

      {/* Jouissance */}
      {(show('destination') || show('article_606')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Refacturation et jouissance</div></div>
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
        {!tree.length
          ? <div className="history-empty">Aucune extraction sauvegardée</div>
          : tree.map(bail => (
            <div key={bail.id}>
              <div className="history-row">
                <button className={`history-btn${bail.id === activeId ? ' active' : ''}`} onClick={() => onSelect(bail)}>
                  <div className="history-name">{bail.data?.immeuble || bail.data?.adresse || bail.file_name}</div>
                  <div className="history-meta">{bail.data?.preneur?.split(',')[0] || '—'} · {formatDate(bail.created_at)}</div>
                </button>
                <button className="history-del" onClick={e => onDelete(bail, e)} title="Supprimer">✕</button>
              </div>
              {bail.avenants?.map(av => (
                <div key={av.id} className="history-row av-row">
                  <button className={`history-btn${av.id === activeId ? ' active' : ''}`} onClick={() => onSelect(av)}>
                    <div className="history-name"><span className="av-tag">A</span>{av.data?.objet_avenant || av.file_name}</div>
                    <div className="history-meta">{formatDate(av.created_at)}</div>
                  </button>
                  <button className="history-del" onClick={e => onDelete(av, e)} title="Supprimer">✕</button>
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
  const [docTypes,     setDocTypes]     = useState([])     // 'bail'|'avenant'|'' per file
  const [fileOrder,    setFileOrder]    = useState([])     // indices ordonnés
  const [detecting,    setDetecting]    = useState(false)  // détection en cours
  const [avenantLinks, setAvenantLinks] = useState({})     // index -> parentId
  const [pertinents,   setPertinents]   = useState([])     // bool per file
  const [raisons,      setRaisons]      = useState([])     // raison non pertinent
  const [linkPhase,    setLinkPhase]    = useState(false)  // phase rattachement post-extraction
  const [extractedMap, setExtractedMap] = useState({})     // index -> {extracted, docType}
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
  function setStatus(i, state, error) { setStatuses(prev => { const n = [...prev]; n[i] = { state, error }; return n }) }

  async function saveExtraction(file, extracted, docType, parentId) {
    const { data: saved } = await supabase.from('extractions')
      .insert({ file_name: file.name, data: extracted, document_type: docType, parent_id: parentId || null })
      .select().single()
    return saved
  }

  // Détection automatique déclenchée au drop
  async function detectFiles(newFiles) {
    setDetecting(true)
    const types      = new Array(newFiles.length).fill('')
    const pertinents = new Array(newFiles.length).fill(null)
    const raisons    = new Array(newFiles.length).fill('')
    const detectData = new Array(newFiles.length).fill(null) // bail_reference data
    const chunks = []
    for (let i = 0; i < newFiles.length; i += 3) chunks.push(newFiles.slice(i, i+3).map((_, j) => i+j))
    for (const chunk of chunks) {
      await Promise.all(chunk.map(async i => {
        try {
          const base64 = await toBase64(newFiles[i])
          const mediaType = getMediaType(newFiles[i])
          const data = await callClaude(base64, mediaType, DETECT_PROMPT)
          types[i]      = data?.type === 'avenant' ? 'avenant' : 'bail'
          pertinents[i] = data?.pertinent !== false
          raisons[i]    = data?.raison || ''
          detectData[i] = { preneur: data?.preneur, bailleur: data?.bailleur, adresse: data?.adresse, immeuble: data?.immeuble }
        } catch (_) { types[i] = 'bail'; pertinents[i] = true }
        setDocTypes([...types])
        setPertinents([...pertinents])
        setRaisons([...raisons])
      }))
    }
    const bailIdx    = types.map((t,i) => t === 'bail'    ? i : -1).filter(i => i >= 0)
    const avenantIdx = types.map((t,i) => t === 'avenant' ? i : -1).filter(i => i >= 0)
    setDocTypes([...types])
    setFileOrder([...bailIdx, ...avenantIdx])
    // Pré-remplir bail lié pour les avenants par matching
    const existingBails = history.filter(h => h.document_type === 'bail')
    const autoLinks = {}
    avenantIdx.forEach(i => {
      if (existingBails.length === 1) {
        autoLinks[i] = existingBails[0].id
      } else if (existingBails.length > 1 && detectData[i]) {
        const match = findBestMatch(detectData[i], existingBails)
        if (match) autoLinks[i] = match.item.id
      }
    })
    setAvenantLinks(autoLinks)
    setDetecting(false)
  }

  function handleFiles(newFiles) {
    const arr = Array.from(newFiles)
    setFiles(arr)
    setDocTypes(arr.map(() => ''))
    setFileOrder(arr.map((_, i) => i))
    setStatuses(arr.map(() => ({})))
    setPertinents(arr.map(() => null))
    setRaisons(arr.map(() => ''))
    setLastError('')
    setAvenantLinks({})
    detectFiles(arr)
  }

  function moveFile(fromIdx, dir) {
    const order = [...fileOrder]
    const pos = order.indexOf(fromIdx)
    const newPos = pos + dir
    if (newPos < 0 || newPos >= order.length) return
    ;[order[pos], order[newPos]] = [order[newPos], order[pos]]
    setFileOrder(order)
  }

  function setDocType(i, type) {
    const n = [...docTypes]; n[i] = type; setDocTypes(n)
    // Réordonner
    const bail2    = n.map((t,x) => t === 'bail'    ? x : -1).filter(x => x >= 0)
    const avenant2 = n.map((t,x) => t === 'avenant' ? x : -1).filter(x => x >= 0)
    setFileOrder([...bail2, ...avenant2])
  }

  async function handleExtract() {
    if (!files.length || loading) return
    setLoading(true)
    setLastError('')
    setStatuses(files.map(() => ({})))
    const order = fileOrder.length ? fileOrder : files.map((_, i) => i)
    const bailIndices    = order.filter(i => (docTypes[i] || 'bail') === 'bail')
    const avenantIndices = order.filter(i => docTypes[i] === 'avenant')
    const availableBails = [...history.filter(h => h.document_type === 'bail')]

    // 1. Extraire les baux d'abord
    for (const i of bailIndices) {
      try {
        setStatus(i, 'loading')
        const base64 = await toBase64(files[i])
        const mediaType = getMediaType(files[i])
        const extracted = await callClaude(base64, mediaType, EXTRACTION_PROMPT)
        const saved = await saveExtraction(files[i], extracted, 'bail', null)
        if (saved) {
          const bwa = { ...saved, avenants: [] }
          availableBails.push(bwa)
          setHistory(prev => [bwa, ...prev])
        }
        setStatus(i, 'done')
      } catch (e) { setStatus(i, 'error', e.message); setLastError(e.message) }
    }

    // 2. Extraire les avenants
    const avenantResults = {}
    for (const i of avenantIndices) {
      try {
        setStatus(i, 'loading')
        const base64 = await toBase64(files[i])
        const mediaType = getMediaType(files[i])
        const extracted = await callClaude(base64, mediaType, AVENANT_PROMPT)
        const match = findBestMatch(extracted?.bail_reference, availableBails)
        avenantResults[i] = { extracted, docType: 'avenant' }
        // Améliorer suggestion si findBestMatch trouve mieux que le choix user
        if (!avenantLinks[i] && match?.item) {
          setAvenantLinks(prev => ({ ...prev, [i]: match.item.id }))
        }
        // Mettre à jour dropdown avec les baux du batch
        setAvenantLinks(prev => {
          if (prev[i]) return prev
          return { ...prev, [i]: match?.item?.id || (availableBails.length === 1 ? availableBails[0].id : null) }
        })
        setStatus(i, 'done')
      } catch (e) { setStatus(i, 'error', e.message); setLastError(e.message) }
    }

    setLoading(false)

    if (Object.keys(avenantResults).length > 0) {
      setExtractedMap(avenantResults)
      setLinkPhase(true)
    }
  }

  async function handleConfirmLinks() {
    let lastSaved = null
    for (const [iStr, { extracted }] of Object.entries(extractedMap)) {
      const i = parseInt(iStr)
      const parentId = avenantLinks[i] || null
      try {
        const saved = await saveExtraction(files[i], extracted, 'avenant', parentId)
        if (saved) {
          lastSaved = saved
          setHistory(prev => parentId
            ? prev.map(b => b.id === parentId ? { ...b, avenants: [...(b.avenants || []), saved] } : b)
            : [saved, ...prev])
        }
      } catch (e) { setLastError(e.message) }
    }
    setLinkPhase(false)
    setExtractedMap({})
    if (lastSaved) setActiveItem(lastSaved)
  }

  async function handleDeleteItem(item, e) {
    e.stopPropagation()
    const label = item.data?.immeuble || item.data?.adresse || item.file_name
    if (!window.confirm(`Supprimer "${label}" ? Cette action est irréversible.`)) return
    await supabase.from('extractions').delete().eq('id', item.id)
    if (activeItem?.id === item.id) setActiveItem(null)
    setHistory(prev => prev.filter(b => b.id !== item.id).map(b => ({ ...b, avenants: (b.avenants || []).filter(a => a.id !== item.id) })))
  }

  async function handleClearHistory() {
    if (!window.confirm("Vider tout l'historique ? Cette action est irréversible.")) return
    await supabase.from('extractions').delete().neq('id', '00000000-0000-0000-0000-000000000000')
    setHistory([]); setHistLoaded(false); setActiveItem(null)
  }

  function handleClear() {
    setFiles([]); setStatuses([]); setActiveItem(null); setDocTypes([])
    setLastError(''); setFileOrder([]); setLinkPhase(false)
    setExtractedMap({}); setAvenantLinks({}); setPertinents([]); setRaisons([])
  }

  const d = activeItem?.data || {}
  const resultTitle = d.immeuble || d.adresse || activeItem?.file_name || ''
  const resultSub = [d.preneur?.split(',')[0], d.bailleur?.split(',')[0], d.date_signature ? `Signé le ${d.date_signature}` : null].filter(Boolean).join(' · ')

  return (
    <>
      <div className="app">
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
            <button className={`nav-item${tab === 'extract' ? ' active' : ''}`} onClick={() => switchTab('extract')}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
              </svg>
              Extraire
            </button>
            <button className={`nav-item${tab === 'history' ? ' active' : ''}`} onClick={() => switchTab('history')}>
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
                  activeItem.document_type === 'avenant' ? activeItem.data?.champs_modifies || {} : activeItem.data || {},
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
            {(!activeItem || linkPhase) && (
              <div className="extract-wrap">

                {/* ── Phase rattachement avenants (post-extraction) ── */}
                {linkPhase && (
                  <div>
                    <div style={{ marginBottom: '16px' }}>
                      <div style={{ fontWeight: 700, fontSize: '15px', marginBottom: '4px' }}>Rattachement des avenants</div>
                      <div style={{ fontSize: '13px', color: 'var(--text2)' }}>
                        Les baux du batch ont été extraits. Vérifiez et confirmez le rattachement de chaque avenant.
                      </div>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                      {Object.entries(extractedMap).map(([iStr, { extracted }]) => {
                        const i = parseInt(iStr)
                        const allBails = history.filter(h => h.document_type === 'bail')
                        return (
                          <div key={i} style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 'var(--r)', padding: '12px 16px' }}>
                            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '12px', flexWrap: 'wrap' }}>
                              <div style={{ minWidth: 0 }}>
                                <div style={{ fontSize: '10px', fontWeight: 700, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.06em', marginBottom: '3px' }}>Avenant</div>
                                <div style={{ fontWeight: 600, fontSize: '13px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{files[i]?.name}</div>
                                {extracted?.bail_reference?.preneur && (
                                  <div style={{ fontSize: '12px', color: 'var(--text2)', marginTop: '2px' }}>
                                    Réf. : {extracted.bail_reference.preneur}{extracted.bail_reference.date_bail_origine ? ` · ${extracted.bail_reference.date_bail_origine}` : ''}
                                  </div>
                                )}
                              </div>
                              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: '4px', flexShrink: 0 }}>
                                <div style={{ fontSize: '10px', fontWeight: 700, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.06em' }}>Bail lié</div>
                                <select
                                  value={avenantLinks[i] || ''}
                                  onChange={e => setAvenantLinks(prev => ({ ...prev, [i]: e.target.value || null }))}
                                  style={{ fontSize: '12px', padding: '5px 10px', borderRadius: '6px', border: '1px solid var(--border2)', background: 'var(--surface)', color: avenantLinks[i] ? 'var(--text)' : 'var(--text3)', cursor: 'pointer', maxWidth: '280px' }}
                                >
                                  <option value="">— Sans rattachement —</option>
                                  {allBails.map(b => (
                                    <option key={b.id} value={b.id}>
                                      {b.data?.immeuble || b.data?.adresse || b.file_name}
                                    </option>
                                  ))}
                                </select>
                              </div>
                            </div>
                          </div>
                        )
                      })}
                    </div>
                    <div style={{ display: 'flex', gap: '8px', marginTop: '16px' }}>
                      <button className="btn primary" onClick={handleConfirmLinks}>Confirmer les rattachements</button>
                      <button className="btn" onClick={() => { setLinkPhase(false); setExtractedMap({}) }}>Annuler</button>
                    </div>
                  </div>
                )}

                {/* ── Queue principale ── */}
                {!linkPhase && (
                  <>
                    <DropZone onFiles={handleFiles} disabled={loading || detecting} />
                    <PageLimitWarning />

                    {files.length > 0 && (
                      <div style={{ marginTop: '10px' }}>
                        {/* En-tête colonnes */}
                        <div style={{ display: 'grid', gridTemplateColumns: '20px 1fr 80px 100px 180px 40px', gap: '8px', padding: '0 4px 6px', borderBottom: '1px solid var(--border)', marginBottom: '4px' }}>
                          <div/>
                          <div style={{ fontSize: '10px', fontWeight: 700, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.06em' }}>Fichier</div>
                          <div style={{ fontSize: '10px', fontWeight: 700, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.06em' }}>Type</div>
                          <div style={{ fontSize: '10px', fontWeight: 700, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.06em' }}>Pertinent</div>
                          <div style={{ fontSize: '10px', fontWeight: 700, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.06em' }}>Bail lié</div>
                          <div/>
                        </div>

                        <div className="file-queue" style={{ marginTop: 0 }}>
                          {(fileOrder.length ? fileOrder : files.map((_, i) => i)).map((fileIdx, pos) => {
                            const f        = files[fileIdx]
                            const st       = statuses[fileIdx] || {}
                            const dt       = docTypes[fileIdx] || ''
                            const isAvenant = dt === 'avenant'
                            const isBail   = dt === 'bail'
                            const pertinent = pertinents[fileIdx]
                            const raison   = raisons[fileIdx] || ''
                            const analyzing = detecting && dt === ''
                            // Baux disponibles = historique + fichiers du batch avec toggle=bail
                            const batchBails = files
                              .map((bf, bi) => docTypes[bi] === 'bail' && bi !== fileIdx ? { id: `batch-${bi}`, file_name: bf.name, _batchIdx: bi } : null)
                              .filter(Boolean)
                            const allBails = [
                              ...history.filter(h => h.document_type === 'bail'),
                              ...batchBails
                            ]
                            return (
                              <div key={fileIdx} className={`queue-item ${st.state || ''}`}
                                style={{ display: 'grid', gridTemplateColumns: '20px 1fr 100px 100px 180px 40px', gap: '8px', alignItems: 'center', padding: '8px 4px', flexWrap: 'nowrap' }}>

                                {/* Ordre ▲▼ */}
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '1px' }}>
                                  <button onClick={() => moveFile(fileIdx, -1)} disabled={pos === 0 || !!st.state}
                                    style={{ background: 'none', border: 'none', color: (pos === 0 || st.state) ? 'var(--border)' : 'var(--text3)', cursor: (pos === 0 || st.state) ? 'default' : 'pointer', padding: 0, fontSize: '9px', lineHeight: 1 }}>▲</button>
                                  <button onClick={() => moveFile(fileIdx, 1)} disabled={pos === fileOrder.length-1 || !!st.state}
                                    style={{ background: 'none', border: 'none', color: (pos === fileOrder.length-1 || st.state) ? 'var(--border)' : 'var(--text3)', cursor: (pos === fileOrder.length-1 || st.state) ? 'default' : 'pointer', padding: 0, fontSize: '9px', lineHeight: 1 }}>▼</button>
                                </div>

                                {/* Nom */}
                                <div style={{ minWidth: 0 }}>
                                  <div style={{ fontWeight: 500, fontSize: '12px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{f.name}</div>
                                  <div style={{ fontSize: '11px', color: 'var(--text3)' }}>{(f.size/1024).toFixed(0)} Ko
                                    {st.state === 'loading' && <span style={{ color: 'var(--accent)', marginLeft: '6px' }}>En cours…</span>}
                                    {st.state === 'done'    && <span style={{ color: 'var(--success)', marginLeft: '6px' }}>✓ Extrait</span>}
                                    {st.state === 'error'   && <span style={{ color: 'var(--danger)', marginLeft: '6px' }} title={st.error}>✕ Erreur</span>}
                                  </div>
                                </div>

                                {/* Toggle Bail/Avenant */}
                                <div>
                                  {analyzing ? (
                                    <span style={{ fontSize: '11px', color: 'var(--text3)', fontStyle: 'italic' }}>Analyse…</span>
                                  ) : (
                                    <div style={{ display: 'flex', gap: 0, border: '1px solid var(--border2)', borderRadius: '6px', overflow: 'hidden', width: 'fit-content' }}>
                                      <button
                                        style={{ padding: '3px 8px', fontSize: '11px', fontWeight: 600, border: 'none', background: isBail ? 'var(--accent)' : 'transparent', color: isBail ? '#fff' : 'var(--text2)', cursor: 'pointer' }}
                                        onClick={() => setDocType(fileIdx, 'bail')}>Bail</button>
                                      <button
                                        style={{ padding: '3px 8px', fontSize: '11px', fontWeight: 600, border: 'none', borderLeft: '1px solid var(--border2)', background: isAvenant ? 'var(--accent)' : 'transparent', color: isAvenant ? '#fff' : 'var(--text2)', cursor: 'pointer' }}
                                        onClick={() => setDocType(fileIdx, 'avenant')}>Avenant</button>
                                    </div>
                                  )}
                                </div>

                                {/* Pertinent */}
                                <div>
                                  {analyzing || pertinent === null ? (
                                    <span style={{ fontSize: '11px', color: 'var(--text3)' }}>—</span>
                                  ) : (
                                    <span title={raison} style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '11px', fontWeight: 600,
                                      padding: '2px 8px', borderRadius: '999px',
                                      background: pertinent ? 'var(--success-bg)' : 'var(--danger-bg)',
                                      color: pertinent ? 'var(--success)' : 'var(--danger)',
                                      cursor: raison ? 'help' : 'default' }}>
                                      {pertinent ? 'Oui' : 'Non'}
                                    </span>
                                  )}
                                </div>

                                {/* Bail lié */}
                                <div>
                                  {isAvenant ? (
                                    <select
                                      value={avenantLinks[fileIdx] || ''}
                                      onChange={e => setAvenantLinks(prev => ({ ...prev, [fileIdx]: e.target.value || null }))}
                                      style={{ fontSize: '11px', padding: '3px 6px', borderRadius: '6px', border: '1px solid var(--border2)', background: 'var(--surface)', color: avenantLinks[fileIdx] ? 'var(--text)' : 'var(--text3)', cursor: 'pointer', width: '100%' }}
                                    >
                                      <option value="">— Bail lié —</option>
                                      {allBails.map(b => (
                                        <option key={b.id} value={b.id}>
                                          {b.data?.immeuble || b.data?.adresse || b.file_name}
                                        </option>
                                      ))}
                                    </select>
                                  ) : <span/>}
                                </div>

                                {/* Supprimer */}
                                <button className="queue-remove" onClick={() => {
                                  setFiles(p => p.filter((_,j) => j !== fileIdx))
                                  setDocTypes(p => p.filter((_,j) => j !== fileIdx))
                                  setStatuses(p => p.filter((_,j) => j !== fileIdx))
                                  setPertinents(p => p.filter((_,j) => j !== fileIdx))
                                  setRaisons(p => p.filter((_,j) => j !== fileIdx))
                                  setFileOrder(fo => fo.filter(x => x !== fileIdx).map(x => x > fileIdx ? x-1 : x))
                                  setAvenantLinks(prev => { const n = {...prev}; delete n[fileIdx]; return n })
                                }}>✕</button>
                              </div>
                            )
                          })}
                        </div>
                      </div>
                    )}

                    {/* Barre d'action */}
                    {files.length > 0 && !loading && !detecting && (
                      <div className="extract-bar">
                        <button className="btn primary" onClick={handleExtract}>
                          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M5 3l14 9-14 9V3z"/></svg>
                          Extraire {files.length > 1 ? `les ${files.length} fichiers` : 'le fichier'}
                        </button>
                        <button className="btn" onClick={handleClear}>Tout effacer</button>
                      </div>
                    )}

                    {detecting && (
                      <div style={{ marginTop: '10px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <div className="progress-track" style={{ flex: 1, margin: 0 }}><div className="progress-bar active" /></div>
                        <span className="status-msg">Analyse en cours…</span>
                      </div>
                    )}

                    {loading && (
                      <div style={{ marginTop: '10px' }}>
                        <div className="progress-track"><div className="progress-bar active" /></div>
                        <div className="status-msg">Extraction en cours…</div>
                      </div>
                    )}

                    {!loading && statuses.some(s => s.state === 'error') && (() => {
                      const done = statuses.filter(s => s.state === 'done').length
                      const errors = statuses.map((s, i) => s.state === 'error' ? { name: files[i]?.name, msg: s.error } : null).filter(Boolean)
                      return (
                        <div style={{ marginTop: '12px', padding: '12px 14px', borderRadius: 'var(--r)', background: 'var(--danger-bg)', border: '1px solid #E8A0A0', fontSize: '12px', color: 'var(--danger)', lineHeight: '1.7' }}>
                          <div style={{ fontWeight: 700, marginBottom: '6px' }}>
                            {done > 0 && <span style={{ color: 'var(--success)', marginRight: '12px' }}>✓ {done} extrait{done > 1 ? 's' : ''}</span>}
                            ✕ {errors.length} erreur{errors.length > 1 ? 's' : ''}
                          </div>
                          {errors.map((e, idx) => (
                            <div key={idx} style={{ display: 'flex', gap: '8px', alignItems: 'flex-start', marginBottom: idx < errors.length - 1 ? '6px' : 0 }}>
                              <span style={{ fontWeight: 600, flexShrink: 0 }}>{e.name}</span>
                              <span>— {e.msg}</span>
                              <button onClick={() => navigator.clipboard.writeText(e.msg)} style={{ marginLeft: 'auto', background: 'none', border: '1px solid #E8A0A0', borderRadius: '4px', color: 'var(--danger)', fontSize: '11px', padding: '1px 6px', cursor: 'pointer', flexShrink: 0 }}>Copier</button>
                            </div>
                          ))}
                        </div>
                      )
                    })()}
                  </>
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
