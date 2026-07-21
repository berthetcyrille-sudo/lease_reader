import { useState, useEffect, useCallback, useRef } from 'react'
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
    { key: 'date_conge', label: 'Date limite de congé' },  // kept for legacy display only
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

const EXTRACTION_PROMPT = `Expert baux commerciaux français. Extrais les données du bail. JSON minifié UNE SEULE LIGNE, sans markdown.

REGLES: Guillemets droits ASCII. Pas de retour a la ligne dans les valeurs. Champs _montant=chiffres bruts sans symbole (ex: 123405.50). null si absent.

CHAMPS:
{"adresse":null,"immeuble":null,"ville":null,"type_bail":null,"duree_totale":null,"duree_ferme":null,"preneur":null,"bailleur":null,"garant":null,"date_effet":null,"date_signature":null,"break_options":[],"notice":null,"date_conge":null,"date_fin":null,"date_limite_travaux":null,"conditions_break":null,"surface_totale_m2":null,"surfaces_detail":[],"parking_nb_places":null,"parking":null,"rie":null,"loyer_signature_montant":null,"loyer_signature":null,"loyer_cours":null,"indexation":null,"indexation_indice":null,"indexation_trimestre_base":null,"indexation_valeur_base":null,"franchise_periodes":[],"franchise":null,"charges":null,"depot_garantie_montant":null,"depot_garantie":null,"travaux_montant":null,"travaux_date_factures":null,"travaux_modalites":null,"participations_travaux":[],"indemnites":[],"indemnites_detail":null,"article_606":null,"conformite":null,"accession":null,"remise_en_etat":null,"maintenance":null,"destination":null,"sous_location":null,"cession":null,"mise_a_disposition":null,"indemnites_restitution":[]}

REGLES PAR CHAMP:
- duree_totale: duree totale du bail (date_effet a date_fin). duree_ferme: si break_options, intervalle date_effet->premiere break option; sinon=duree_totale; si mentionne explicitement, utiliser cette valeur.
- surfaces_detail: [{\"categorie\":\"Bureaux\",\"niveau\":\"5eme etage\",\"surface_m2\":\"2224.98\",\"prix_unitaire\":\"290\",\"loyer_annuel\":\"645244\"}]. categorie JAMAIS null: etage/plateau->Bureaux, sous-sol/emplacement/lot numerote->Stationnement, exterieur->Stationnement, doute->Bureaux.
- notice: DUREE du préavis pour donner congé, exprimée en mois uniquement (ex: "6 mois", "3 mois"). NE PAS mettre une date. Si le bail dit "au moins six (6) mois avant la date d'échéance" → notice="6 mois".
- mise_a_disposition: si le bail prevoit une mise a disposition anticipee des locaux (avant la date d'effet officielle du bail). Format: {"date_debut":"jj/mm/aaaa","date_fin":"jj/mm/aaaa","loyer_paye":"Oui/Non/Partiel","charges_payees":"Oui/Non/Partiel","conditions":"texte libre des conditions financieres pendant cette periode"}. null si aucune mise a disposition anticipee.
- break_options: liste COMPLETE et EXHAUSTIVE de toutes les dates auxquelles le PRENEUR peut effectivement sortir avant le terme. Format: ["31/08/2028","31/08/2029","31/08/2030"]. REGLE CRITIQUE: les Conditions Particulières (CP) priment TOUJOURS sur les Conditions Générales (CG). Si les CG preVoient des echeances triennales MAIS que les CP contiennent une renonciation expresse du preneur ("le PRENEUR renonce expressement a sa faculte de resiliation triennale" ou equivalent) → il N'Y A PAS de break triennale, indiquer uniquement les dates expressement stipulees dans les CP. REGLE DE CALCUL quand les breaks sont confirmees par les CP: "a l'expiration de chaque periode triennale" → date_effet + 3 ans, + 6 ans (si < date_fin). "a l'expiration de la Neme annee" → date_effet + N ans. Inclure TOUTES ces dates calculees meme si non ecrites explicitement. Ne PAS inclure date_fin.
- loyer_signature_montant: MONTANT ANNUEL TOTAL HT/HC. JAMAIS prix unitaire/m². Si tableau par lot: additionner les loyer_annuel. INTERDIT de retourner null si un loyer figure dans le document.
- loyer_cours: loyer annuel "de base" au sens indexation. Identique a loyer_signature_montant sauf mention contraire. JAMAIS prix unitaire/m².
- indexation_indice: code de l indice parmi: "ILAT","ILC","ICC","IRL","IPC","BT01","AUTRE". null si non renseigne.
- indexation_trimestre_base: trimestre de reference si EXPLICITEMENT indique dans le bail (ex: "3T2025"). null si le bail dit "dernier indice publie a la date de signature" sans preciser lequel.
- indexation_valeur_base: valeur numerique de l indice si EXPLICITEMENT mentionnee (ex: "120.5"). null si non mentionnee.
- franchise_periodes: TOUTES les franchises, y compris conditionnelles. [{\"date_debut\":\"jj/mm/aaaa\",\"date_fin\":\"jj/mm/aaaa\",\"duree\":\"6 mois\",\"montant\":\"123405\",\"surface_assiette\":\"LC1 (701 m²)\",\"indexation_incluse\":\"Non\",\"condition\":null}]. montant=chiffres bruts (calcule si non explicite: loyer_annuel_assiette*duree_mois/12). condition=texte si conditionnelle, null sinon.
- participations_travaux: UNIQUEMENT si le bail prevoit une enveloppe financiere DISTINCTE de la franchise, specifiquement dediee aux travaux (ex: "le BAILLEUR verse X euros pour les travaux" avec un calendrier de facturation propre). EXCLURE: les franchises de loyer qualifiees de participation aux travaux (ex: "franchise accordee au titre de la participation aux travaux") — ces franchises doivent figurer UNIQUEMENT dans franchise_periodes. En cas de doublon franchise/travaux sur le meme montant, privilegier franchise_periodes. Format: [{\"libelle\":\"denomination exacte\",\"montant\":\"822701\",\"date_limite\":\"31/12/2024\",\"remarque\":null}]. libelle OBLIGATOIRE.
- parking_nb_places: ex: "114 places (98 interieures + 16 exterieures)"
- indemnites: UNIQUEMENT indemnites liees a une option (break, renouvellement, fin de bail). EXCLURE: honoraires, cautionnements, penalites. [{\"motif\":\"...\",\"due_par\":\"Preneur ou Bailleur\",\"montant\":\"chiffres bruts\",\"date_limite\":\"...\"}]`

const AVENANT_PROMPT = `Expert baux commerciaux français. Ce document est un AVENANT. JSON minifié UNE SEULE LIGNE, sans markdown.

REGLES: Guillemets droits ASCII. Champs _montant=chiffres bruts. Dans champs_modifies: null pour les champs NON modifies par l'avenant.

surface_change_type: "inchangee"/"ajout"/"retrait"/"substitution"/"mixte".
surfaces_delta: surfaces UNIQUEMENT concernees par la modif (ajoutees ou retirees). Ajouter "sens":"ajout" ou "sens":"retrait". categorie JAMAIS null.
surfaces_avant: tableau EXACT des surfaces telles qu'elles etaient AVANT cet avenant, tel que decrit dans le bail d'origine mentionne dans ce document. categorie JAMAIS null. null si surface_change_type="inchangee".
surfaces_apres: tableau EXACT des surfaces APRES cet avenant. REGLE STRICTE: regrouper par categorie si plusieurs lignes de meme categorie (ex: 2 lignes Bureaux → une seule ligne avec la surface totale). NE PAS INVENTER de lignes. NE PAS dupliquer. La surface totale de surfaces_apres doit etre egale a surface_totale_m2. categorie JAMAIS null. null si surface_change_type="inchangee".

{"bail_reference":{"preneur":null,"bailleur":null,"date_bail_origine":null,"adresse":null,"immeuble":null},"date_effet_avenant":null,"date_signature_avenant":null,"objet_avenant":null,"surface_change_type":"inchangee","surfaces_delta":null,"surfaces_avant":null,"surfaces_apres":null,"champs_modifies":{"adresse":null,"immeuble":null,"ville":null,"type_bail":null,"duree_totale":null,"duree_ferme":null,"preneur":null,"bailleur":null,"garant":null,"date_effet":null,"date_signature":null,"break_options":null,"notice":null,"date_conge":null,"date_fin":null,"date_limite_travaux":null,"conditions_break":null,"surface_totale_m2":null,"surfaces_detail":null,"parking_nb_places":null,"parking":null,"rie":null,"loyer_signature_montant":null,"loyer_signature":null,"loyer_cours":null,"indexation":null,"franchise_periodes":null,"franchise":null,"charges":null,"depot_garantie_montant":null,"depot_garantie":null,"travaux_montant":null,"travaux_date_factures":null,"travaux_modalites":null,"participations_travaux":null,"indemnites":null,"indemnites_detail":null,"article_606":null,"conformite":null,"accession":null,"remise_en_etat":null,"maintenance":null,"destination":null,"sous_location":null,"cession":null,"mise_a_disposition":null,"indemnites_restitution":[]}}

REGLES PAR CHAMP (champs_modifies):
- loyer_signature_montant: montant annuel total HT/HC. null si non modifie. JAMAIS prix unitaire/m².
- franchise_periodes: TOUTES les nouvelles franchises de l'avenant. [{\"date_debut\":\"jj/mm/aaaa\",\"date_fin\":\"jj/mm/aaaa\",\"duree\":\"6 mois\",\"montant\":\"123405\",\"surface_assiette\":\"LC1 (701 m²)\",\"indexation_incluse\":\"Non\",\"condition\":null}]. null si aucune franchise dans l'avenant.
- participations_travaux: UNIQUEMENT si enveloppe financiere DISTINCTE de la franchise, dediee aux travaux avec calendrier de facturation propre. Ne JAMAIS y mettre une franchise de loyer meme si qualifiee "au titre des travaux" — celle-ci va dans franchise_periodes. En cas de doute sur meme montant, privilegier franchise_periodes. Format: [{\"libelle\":\"denomination exacte\",\"montant\":\"822701\",\"date_limite\":\"31/12/2024\",\"remarque\":null}]. null si non concerne.
- surfaces_detail: tableau complet post-avenant UNIQUEMENT si l'avenant redefinit completement l'assiette. null sinon (utiliser surfaces_apres a la place).`

const DETECT_PROMPT = `Analyse ce document. Le nom du fichier est un indice important. Reponds UNIQUEMENT avec ce JSON sur une ligne:
{"type":"bail","pertinent":true,"raison":"","preneur":"","bailleur":"","adresse":"","immeuble":""}
Regles strictes:
- pertinent: true UNIQUEMENT si le document est un bail commercial original ou un avenant a un bail commercial (y compris s'il est intitule "protocole" ou "accord" mais qu'il modifie les conditions d'un bail: loyer, franchise, duree, surfaces). false dans TOUS les autres cas: side letter TVA, courrier simple, facture, plan, etat des lieux, diagnostic energetique, acte de cautionnement autonome, police d'assurance, mandat de gestion, proces-verbal d'assemblee, ou tout document qui ne modifie pas lui-meme les conditions d'un bail commercial
- type: "bail" si bail commercial original, "avenant" si avenant/rectificatif/protocole modificatif d'un bail (reduction de loyer, franchise supplementaire, changement de surface, etc.)
- raison: explication courte si pertinent:false (ex: "side letter TVA", "etat des lieux", "diagnostic energetique")
- preneur, bailleur, adresse, immeuble: extrais ces valeurs du document pour identifier le bail associe`

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

function normalizeDate(val) {
  if (!val) return null
  const v = String(val)
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(v)) return v
  const iso = v.match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`
  const months = { janvier:1,février:2,mars:3,avril:4,mai:5,juin:6,juillet:7,août:8,septembre:9,octobre:10,novembre:11,décembre:12 }
  const cleaned = v.toLowerCase().replace(/1er/, '1').replace(/[èe]me/, '')
  const fr = cleaned.match(/(\d{1,2})\s+([a-zéû]+)\s+(\d{4})/)
  if (fr && months[fr[2]]) return `${String(parseInt(fr[1])).padStart(2,'0')}/${String(months[fr[2]]).padStart(2,'0')}/${fr[3]}`
  return v
}

function formatDate(iso) {
  if (!iso) return '—'
  return new Date(iso).toLocaleDateString('fr-FR', { day: '2-digit', month: 'short', year: 'numeric' })
}

// Parse a raw montant string to a float number (strips currency symbols, spaces)
// Convert any value to a renderable string — prevents React error #31 when Claude returns objects
function safeStr(val) {
  if (val === null || val === undefined) return null
  if (typeof val === 'string') return val || null
  if (typeof val === 'number' || typeof val === 'boolean') return String(val)
  if (Array.isArray(val)) return val.map(safeStr).filter(Boolean).join(', ') || null
  if (typeof val === 'object') {
    // Try common text keys first
    const txt = val.commentaire || val.texte || val.valeur || val.value || val.text || val.description
    if (txt) return safeStr(txt)
    // Fallback: join all string values
    return Object.entries(val).map(([k, v]) => `${k}: ${safeStr(v)}`).join(' · ') || null
  }
  return String(val)
}

function parseAmount(val) {
  if (!val) return null
  const n = parseFloat(String(val).replace(/[^0-9.,]/g, '').replace(',', '.'))
  return isNaN(n) ? null : n
}

// Condense verbose parking text to short summary e.g. "99 int. + 30 ext. = 129 places"
function parseParkingShort(val) {
  if (!val) return null
  // Handle object: try common keys
  if (typeof val === 'object' && !Array.isArray(val)) {
    val = val.total || val.nombre || val.nb || val.value || val.texte || val.commentaire || Object.values(val).find(v => v) || ''
  }
  const s = String(val)
  if (!s || s === 'null' || s === 'undefined') return null
  const intMatch = s.match(/(\d+)\s+int[eé]r/i)
  const extMatch = s.match(/(\d+)\s+ext[eé]r/i)
  const totalMatch = s.match(/^(\d+)\s+place/i)
  if (intMatch || extMatch) {
    const nb_int = intMatch ? parseInt(intMatch[1]) : 0
    const nb_ext = extMatch ? parseInt(extMatch[1]) : 0
    const total = nb_int + nb_ext
    const parts = []
    if (nb_int) parts.push(`${nb_int} int.`)
    if (nb_ext) parts.push(`${nb_ext} ext.`)
    return parts.join(' + ') + (total ? ` = ${total} places` : '')
  }
  if (totalMatch) return s.match(/^\d+\s+places?[^(]*/i)?.[0]?.trim() || s
  return s
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

const MAX_BREAKS    = 6
const MAX_FRANCHISE = 8
const MAX_INDEM     = 5
const MAX_SURF      = 8
const MAX_TRAV      = 4
const MAX_PALIERS   = 4
const MAX_ABAT      = 4
const MAX_IB        = 4

function buildExcelHeaders() {
  const breakCols     = Array.from({ length: MAX_BREAKS },    (_, i) => `Break option ${i+1}`)
  const franchiseCols = Array.from({ length: MAX_FRANCHISE }, (_, i) => [
    `Franchise P${i+1} - Debut`, `Franchise P${i+1} - Fin`, `Franchise P${i+1} - Duree`,
    `Franchise P${i+1} - Assiette`, `Franchise P${i+1} - Montant`, `Franchise P${i+1} - Indexation`, `Franchise P${i+1} - Condition`,
  ]).flat()
  const indemnCols    = Array.from({ length: MAX_INDEM },     (_, i) => [
    `Indemnite ${i+1} - Motif`, `Indemnite ${i+1} - Due par`, `Indemnite ${i+1} - Montant`, `Indemnite ${i+1} - Echeance`,
  ]).flat()
  const surfCols      = Array.from({ length: MAX_SURF },      (_, i) => [
    `Surface ${i+1} - Categorie`, `Surface ${i+1} - Niveau`, `Surface ${i+1} - m2`, `Surface ${i+1} - Prix m2`, `Surface ${i+1} - Loyer/an`,
  ]).flat()
  const travCols      = Array.from({ length: MAX_TRAV },      (_, i) => [
    `Travaux ${i+1} - Libelle`, `Travaux ${i+1} - Montant`, `Travaux ${i+1} - Date limite`,
  ]).flat()
  const palierCols    = Array.from({ length: MAX_PALIERS },   (_, i) => [
    `Palier ${i+1} - Debut`, `Palier ${i+1} - Fin`, `Palier ${i+1} - Montant annuel`, `Palier ${i+1} - Description`,
  ]).flat()
  const abatCols      = Array.from({ length: MAX_ABAT },      (_, i) => [
    `Abattement ${i+1} - Debut`, `Abattement ${i+1} - Fin`, `Abattement ${i+1} - Montant annuel`, `Abattement ${i+1} - Description`,
  ]).flat()
  const ibCols        = Array.from({ length: MAX_IB },        (_, i) => [
    `Indem.break ${i+1} - Date break`, `Indem.break ${i+1} - Motif`, `Indem.break ${i+1} - Montant`, `Indem.break ${i+1} - Formule`,
  ]).flat()
  return [
    'ID', 'Bail lié (ID)',
    'Type', 'Actif / Immeuble', 'Adresse', 'Ville',
    'Preneur', 'Bailleur',
    'Type de bail', 'Duree totale', 'Duree ferme',
    'Date effet', 'Date signature', 'Date fin', 'Date conge limite', 'Preavis', 'Date limite travaux preneur',
    ...breakCols,
    'Conditions break',
    'Surface totale m2', 'Parking nb places', 'Parking loyer unitaire (€/place/an)', 'RIE',
    ...surfCols,
    'Loyer HT/HC annuel signature', 'Loyer de base annuel', 'Indexation', 'Indice base - Code', 'Indice base - Trimestre', 'Indice base - Valeur', 'Indice base - Source', 'Loyer signature detail',
    ...franchiseCols,
    'Franchise modalites',
    'Charges TEOM',
    'Depot garantie montant', 'Depot garantie modalites',
    'Travaux montant unique', 'Travaux date limite', 'Travaux modalites',
    ...travCols,
    ...indemnCols,
    'Article 606', 'Conformite', 'Remise en etat', 'Sous-location', 'Cession', 'Destination', 'Maintenance', 'Accession',
    // Loyer variable
    'Loyer variable - Type', 'Loyer variable - Taux', 'Loyer variable - Assiette', 'Loyer variable - Plancher', 'Loyer variable - Plafond', 'Loyer variable - Formule',
    ...palierCols,
    ...abatCols,
    ...ibCols,
    // Indemnites restitution
    'Indem.restitution 1 - Terme', 'Indem.restitution 1 - Due par', 'Indem.restitution 1 - Motif', 'Indem.restitution 1 - Montant', 'Indem.restitution 1 - Calcul',
    'Indem.restitution 2 - Terme', 'Indem.restitution 2 - Due par', 'Indem.restitution 2 - Motif', 'Indem.restitution 2 - Montant', 'Indem.restitution 2 - Calcul',
    'Indem.restitution 3 - Terme', 'Indem.restitution 3 - Due par', 'Indem.restitution 3 - Motif', 'Indem.restitution 3 - Montant', 'Indem.restitution 3 - Calcul',
    // Avenant-specific
    'Objet avenant', 'Date effet avenant', 'Date signature avenant', 'Bail lie', 'Modif surfaces type',
  ]
}

function buildExcelRow(item, bailParentName, bailParentData) {
  const isAv   = item.document_type === 'avenant'
  const raw    = item.data || {}
  const mods   = isAv ? (raw.champs_modifies || {}) : {}
  const meta   = isAv ? raw : {}
  // For avenants: merge bail parent data with champs_modifies (non-null overrides base)
  const base   = isAv ? (bailParentData || {}) : raw
  const d      = isAv
    ? Object.fromEntries(
        [...new Set([...Object.keys(base), ...Object.keys(mods)])].map(k => [
          k, (mods[k] !== null && mods[k] !== undefined && !(Array.isArray(mods[k]) && mods[k].length === 0))
             ? mods[k] : base[k]
        ])
      )
    : raw

  const v    = (val) => { const s = safeStr(val); return s || '' }
  const amt  = (val) => { const n = parseAmount(val); return n !== null ? n : '' }
  const surf = (val) => { const n = parseFloat(String(val || '').replace(',', '.')); return isNaN(n) ? '' : n }

  const breaks    = Array.isArray(d.break_options)          ? d.break_options          : []
  const franchise = Array.isArray(d.franchise_periodes)     ? d.franchise_periodes     : []
  const indem     = Array.isArray(d.indemnites)             ? d.indemnites             : []
  const surfaces  = Array.isArray(d.surfaces_detail)        ? d.surfaces_detail        : []
  const trav      = Array.isArray(d.participations_travaux) ? d.participations_travaux : []
  const paliers   = Array.isArray(d.paliers_loyer)          ? d.paliers_loyer          : []
  const abats     = Array.isArray(d.abattements)            ? d.abattements            : []
  const ibs       = Array.isArray(d.indemnites_break)       ? d.indemnites_break       : []

  const breakVals = Array.from({ length: MAX_BREAKS },    (_, i) => v(breaks[i]) )
  const fracVals  = Array.from({ length: MAX_FRANCHISE }, (_, i) => [
    v(franchise[i]?.date_debut), v(franchise[i]?.date_fin), v(franchise[i]?.duree),
    v(franchise[i]?.surface_assiette), amt(franchise[i]?.montant), v(franchise[i]?.indexation_incluse), v(franchise[i]?.condition),
  ]).flat()
  const indemVals = Array.from({ length: MAX_INDEM },     (_, i) => [
    v(indem[i]?.motif), v(indem[i]?.due_par), amt(indem[i]?.montant), v(indem[i]?.date_limite),
  ]).flat()
  // Enrich surfaces with computed prix_unitaire
  const surfEnriched = computeUnitPrices(surfaces, d.parking_nb_places, null)
  const surfVals  = Array.from({ length: MAX_SURF },      (_, i) => {
    const r = surfEnriched[i]
    if (!r) return ['', '', '', '', '']
    const up = r.prix_unitaire ? amt(r.prix_unitaire) :
               (() => { const l = parseAmount(r.loyer_annuel); const s = parseFloat(String(r.surface_m2||'').replace(',','.'))||0; return (l!==null&&s>0) ? Math.round(l/s) : '' })()
    return [v(r.categorie), v(r.niveau), surf(r.surface_m2), up, amt(r.loyer_annuel)]
  }).flat()
  // Parking unit price
  const pkUnit = computeParkingUnitPrice(d.parking_nb_places, surfaces)
  const travVals  = Array.from({ length: MAX_TRAV },      (_, i) => [
    v(trav[i]?.libelle), amt(trav[i]?.montant), v(trav[i]?.date_limite),
  ]).flat()
  const palierVals = Array.from({ length: MAX_PALIERS },   (_, i) => [
    v(paliers[i]?.date_debut), v(paliers[i]?.date_fin), amt(paliers[i]?.montant), v(paliers[i]?.description),
  ]).flat()
  const abatVals   = Array.from({ length: MAX_ABAT },      (_, i) => [
    v(abats[i]?.date_debut), v(abats[i]?.date_fin), amt(abats[i]?.montant_annuel), v(abats[i]?.description),
  ]).flat()
  const ibVals     = Array.from({ length: MAX_IB },        (_, i) => [
    v(ibs[i]?.break_date), v(ibs[i]?.motif), amt(ibs[i]?.montant), v(ibs[i]?.calcul),
  ]).flat()

  return [
    (item.id || '').slice(0, 8),
    isAv ? (item.parent_id || '').slice(0, 8) : '',
    isAv ? 'Avenant' : 'Bail',
    v(d.immeuble || raw.bail_reference?.immeuble),
    v(d.adresse  || raw.bail_reference?.adresse),
    v(d.ville),
    v(shortPartyName(d.preneur  || raw.bail_reference?.preneur)),
    v(shortPartyName(d.bailleur || raw.bail_reference?.bailleur)),
    v(d.type_bail), v(d.duree_totale), v(d.duree_ferme),
    // Date effet / signature : pour avenant, utiliser les dates propres à l'avenant
    isAv ? v(meta.date_effet_avenant) : v(d.date_effet),
    isAv ? v(meta.date_signature_avenant) : v(d.date_signature),
    v(d.date_fin), v(d.date_conge), v(d.notice), v(d.date_limite_travaux),
    ...breakVals,
    v(d.conditions_break),
    surf(d.surface_totale_m2), parseParkingShort(d.parking_nb_places) || '', pkUnit || '', v(d.rie),
    ...surfVals,
    amt(d.loyer_signature_montant), amt(d.loyer_cours), v(d.indexation),
    // Indice base
    (() => {
      const indice = d.indexation_indice
      if (!indice) return ['', '', '', '']
      if (d.indexation_valeur_base) return [indice, d.indexation_trimestre_base || '', d.indexation_valeur_base, 'bail']
      // Compute from static table synchronously
      const table = INSEE_STATIC[indice.toUpperCase()]
      if (!table) return [indice, '', '', '']
      const dateStr = d.date_signature || d.date_effet
      const m = String(dateStr || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/)
      if (!m) return [indice, '', '', '']
      const month = parseInt(m[2]), year = parseInt(m[3]), q = Math.ceil(month / 3)
      const targetY = year, targetQ = q
      const candidates = table.filter(row => {
        const rm = row.q.match(/(\d)T(\d{4})/); if (!rm) return false
        const y = parseInt(rm[2]), qq = parseInt(rm[1])
        return y < targetY || (y === targetY && qq <= targetQ)
      })
      if (!candidates.length) return [indice, '', '', 'INSEE']
      const last = candidates[candidates.length - 1]
      return [indice, last.q, last.v, 'INSEE (table)']
    })(),
    v(d.loyer_signature),
    ...fracVals,
    v(d.franchise), v(d.charges),
    amt(d.depot_garantie_montant), v(d.depot_garantie),
    amt(d.travaux_montant), v(d.travaux_date_factures), v(d.travaux_modalites),
    ...travVals,
    ...indemVals,
    v(d.article_606), v(d.conformite), v(d.remise_en_etat), v(d.sous_location), v(d.cession), v(d.destination), v(d.maintenance), v(d.accession),
    // Loyer variable
    v(d.loyer_variable?.type), v(d.loyer_variable?.taux), v(d.loyer_variable?.assiette),
    amt(d.loyer_variable?.plancher), amt(d.loyer_variable?.plafond), v(d.loyer_variable?.regles),
    ...palierVals,
    ...abatVals,
    ...ibVals,
    // Indemnites restitution
    ...Array.from({ length: 3 }, (_, i) => {
      const r = (d.indemnites_restitution || [])[i]
      if (!r) return ['', '', '', '', '']
      return [v(r.terme), v(r.due_par), v(r.motif), amt(r.montant), v(r.calcul)]
    }).flat(),
    // Avenant-specific
    v(meta.objet_avenant), v(meta.date_effet_avenant), v(meta.date_signature_avenant),
    bailParentName || '', v(meta.surface_change_type),
  ]
}
function exportToExcel(items, fileName) {
  // items: array of {item, parentName} OR single item (legacy)
  let rows
  if (Array.isArray(items)) {
    rows = items.map(({ item, parentName, parentData }) => buildExcelRow(item, parentName, parentData))
  } else {
    // legacy single call: items is actually a data object
    const fakeItem = { document_type: 'bail', data: items, file_name: fileName }
    rows = [buildExcelRow(fakeItem, '')]
  }

  const headers = buildExcelHeaders()
  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows])

  // Column widths
  // Bold header row
  headers.forEach((_, colIdx) => {
    const addr = XLSX.utils.encode_cell({ r: 0, c: colIdx })
    if (ws[addr]) ws[addr].s = { font: { bold: true } }
  })

  ws['!cols'] = headers.map(h => ({
    wch: h.includes('detail') || h.includes('modalites') || h.includes('Condition') || h.includes('Motif') ? 40
       : h.includes('Libelle') || h.includes('Assiette') ? 35
       : h.includes('Preneur') || h.includes('Bailleur') || h.includes('Objet') ? 30
       : 18
  }))

  // Number format on amount columns
  headers.forEach((h, colIdx) => {
    if (!h.includes('Montant') && !h.includes('montant') && !h.includes('Loyer') && !h.includes('loyer') && !h.includes('m2') && !h.includes('Prix')) return
    rows.forEach((_, rowIdx) => {
      const cell = ws[XLSX.utils.encode_cell({ r: rowIdx + 1, c: colIdx })]
      if (cell && typeof cell.v === 'number') { cell.t = 'n'; cell.z = '#,##0' }
    })
  })

  // Header row style (bold via sheetjs-style not available, use freeze pane instead)
  ws['!freeze'] = { xSplit: 0, ySplit: 1 }

  XLSX.utils.book_append_sheet(wb, ws, 'Base de données')
  const safeName = (fileName || 'lease_abstract').replace(/\.[^.]+$/, '').replace(/[^a-zA-Z0-9_\-]/g, '_')
  // Use write + blob to support cell styles
  try {
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true })
    const blob = new Blob([wbout], { type: 'application/octet-stream' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a'); a.href = url; a.download = `${safeName}.xlsx`
    document.body.appendChild(a); a.click()
    setTimeout(() => { URL.revokeObjectURL(url); document.body.removeChild(a) }, 100)
  } catch {
    XLSX.writeFile(wb, `${safeName}.xlsx`)
  }
}

function exportAllToExcel(tree) {
  const rows = []
  tree.forEach(bail => {
    const parentName = bail.data?.immeuble || bail.data?.adresse || bail.file_name
    const parentData = bail.data || {}
    rows.push({ item: bail, parentName: '', parentData: null })
    const sortedAv = [...(bail.avenants || [])].sort((a, b) => {
      const toS = d => { const m = String(d||'').match(/^(\d{2})\/(\d{2})\/(\d{4})$/); return m ? `${m[3]}-${m[2]}-${m[1]}` : String(d||'') }
      return toS(a.data?.date_effet_avenant || a.data?.date_signature_avenant || a.created_at)
            .localeCompare(toS(b.data?.date_effet_avenant || b.data?.date_signature_avenant || b.created_at))
    })
    sortedAv.forEach(av => rows.push({ item: av, parentName, parentData }))
  })
  exportToExcel(rows, 'lease_abstract_complet')
}

const BREAK_PROMPT = `Expert baux commerciaux français. Analyse UNIQUEMENT la clause de durée et de résiliation de ce bail. Retourne UNIQUEMENT un JSON minifié sur UNE SEULE LIGNE : {"date_effet":"jj/mm/aaaa","date_fin":"jj/mm/aaaa","break_options":["jj/mm/aaaa",...]}

REGLE ABSOLUE pour break_options : liste COMPLETE et EXHAUSTIVE.
REGLE CRITIQUE CP PRIMENT SUR CG: si les Conditions Générales preVoient des echeances triennales MAIS que les Conditions Particulières contiennent une renonciation expresse du preneur ("le PRENEUR renonce expressement a sa faculte de resiliation triennale" ou equivalent) → PAS de break triennale, indiquer UNIQUEMENT les dates expressement stipulees dans les CP.
Si breaks confirmees:
- "a l'expiration de chaque periode triennale" → date_effet + 3 ans, + 6 ans (si < date_fin)
- "a l'expiration de la Neme annee" → date_effet + N ans
- Inclure TOUTES ces dates calculees meme si non ecrites explicitement
- Trier chronologiquement. Ne PAS inclure date_fin.
- Exemple avec triennales + 4eme + 5eme: bail 01/09/2025→31/08/2034 → ["31/08/2028","31/08/2029","31/08/2030","31/08/2031"]
- Exemple avec renonciation aux triennales: bail 31/12/2023→14/04/2033, CP stipulent "congé possible pour le 31/08/2030 uniquement" → ["31/08/2030"]`

const FINANCIAL_PROMPT = `Expert baux commerciaux français. Extrais UNIQUEMENT les données financières critiques de ce bail ou avenant. JSON minifié UNE SEULE LIGNE, sans markdown. Guillemets droits ASCII. Montants=chiffres bruts sans symbole.

{"loyer_signature_montant":null,"loyer_signature":null,"paliers_loyer":[],"abattements":[],"loyer_variable":null,"franchise_periodes":[],"participations_travaux":[],"indemnites_break":[]}

REGLES PAR CHAMP:

loyer_signature_montant: MONTANT ANNUEL TOTAL HT/HC. Jamais prix unitaire/m². Si tableau par lot: additionner tous les loyer_annuel. INTERDIT de retourner null si un loyer figure dans le document.

loyer_signature: texte descriptif complet du loyer (detail par composante, prix unitaires, etc.)

paliers_loyer: tableau si le loyer evolue par etapes a des dates definies (ex: loyer annuel reduit pendant N mois puis loyer plein). Format: [{"date_debut":"jj/mm/aaaa","date_fin":"jj/mm/aaaa","montant":"123456","description":"ex: loyer reduit periode travaux"}]. [] si aucun palier.

abattements: tableau de toutes les reductions temporaires de loyer (ex: abattement RIE, reduction pendant franchise partielle, loyer minoré conditionnel). Format: [{"date_debut":"jj/mm/aaaa","date_fin":"jj/mm/aaaa","montant_annuel":"12345","description":"ex: reduction RIE jusqu a mise en service"}]. [] si aucun abattement.

loyer_variable: si le bail contient une clause de loyer variable ou indexe sur le CA/chiffre d affaires. Format: {"type":"CA ou autre","taux":"ex: 3%","assiette":"ex: CA TTC annuel","plancher":"montant brut ou null","plafond":"montant brut ou null","regles":"texte complet de la formule et des conditions de declenchement"}. null si pas de loyer variable.

franchise_periodes: TOUTES les franchises SANS EXCEPTION, y compris conditionnelles et complementaires. Format: [{"date_debut":"jj/mm/aaaa","date_fin":"jj/mm/aaaa","duree":"6 mois","montant":"123405","surface_assiette":"ex: LC1 (701 m²)","indexation_incluse":"Non","condition":"null ou texte si conditionnelle ex: si non-delivrance de conge au 31/08/2030"}]. montant: calculer si non explicite (loyer_annuel_assiette * duree_mois / 12).

participations_travaux: TOUTES les enveloppes de participation financiere du bailleur aux travaux du preneur. Format: [{"libelle":"denomination exacte ex: Locaux Initiaux R+5","montant":"822701","date_limite":"31/12/2024","remarque":null}]. libelle OBLIGATOIRE, jamais null.

indemnites_restitution: tableau des indemnites FORFAITAIRES DE REMISE EN ETAT contractuellement prevues, dont le montant varie selon la date de sortie du preneur (break ou terme). Ce champ cible UNIQUEMENT les clauses du type "le PRENEUR versera au BAILLEUR une indemnite forfaitaire de X euros en cas de depart a compter de la Neme annee". Creer UNE LIGNE PAR TRANCHE DE MONTANT. EXCLURE absolument: penalites d immobilisation pour non-restitution tardive, indemnites d eviction, depots de garantie, penalites de retard, indemnites de non-renouvellement. REGLE terme: calculer la date exacte depuis date_effet + N annees (ex: 29/06/2026 + 6 ans = 28/06/2032). REGLE calcul: indiquer l indexation si prevue. Format: [{"terme":"28/06/2032 (expiration 6eme annee)","due_par":"Preneur","motif":"Indemnite forfaitaire remise en etat","montant":"115840","calcul":"Forfait indexe BT01: indemnite revisee = base x (indice revision / indice base)"}]. null si aucune indemnite forfaitaire de ce type.
indemnites_break: Sommes dues par le PRENEUR au BAILLEUR UNIQUEMENT en cas d exercice de son droit de CONGE (sortie anticipee par le preneur via l option de break). TROIS CAS A DISTINGUER:
1) FORFAIT CHIFFRE PAR DATE DE BREAK: montant ou formule specifique par echéance (ex: "6 mois de loyer si conge au 31/08/2028") -> une ligne par break date.
2) INDEMNITE DE RESTITUTION/REMISE EN ETAT: si le bail prevoit une indemnite forfaitaire due a la restitution en cas de sortie anticipee (ex: "indemnite forfaitaire de remise en etat de 115 840 € en cas de depart a compter de la 6eme annee") -> inclure avec break_date=date de la premiere break concernee et calcul=formule ou texte.
3) REMBOURSEMENT DES MESURES D ACCOMPAGNEMENT SI CONGE: clause generale de remboursement des avantages (franchises, MDA, travaux) si le preneur exerce son conge avant terme -> une ligne sans break_date specifique.
A EXCLURE de indemnites_break: les clauses de remboursement uniquement en cas de CESSION du bail ou du fonds (pas un conge). Ces clauses-la relèvent de la cession et ne sont pas des indemnites de break.
NE PAS INVENTER de montants. Format: [{"break_date":"31/08/2028 ou null","motif":"texte","montant":"chiffres bruts ou null","calcul":"formule ou texte de la clause"}].`

// ─── JSON cleaning & parsing ──────────────────────────────────────────────────

function ensureArray(val) {
  if (!val) return null
  if (Array.isArray(val)) return val
  if (typeof val === 'string') {
    try { const p = JSON.parse(val); return Array.isArray(p) ? p : [val] } catch (_) { return [val] }
  }
  return null
}

// Normalize categorie values to canonical set
const CAT_MAP = {
  'bureaux': 'Bureaux', 'bureau': 'Bureaux', 'office': 'Bureaux', 'open space': 'Bureaux', 'plateau': 'Bureaux',
  'stationnement': 'Stationnement', 'parking': 'Stationnement', 'parking_interieur': 'Stationnement',
  'parking_exterieur': 'Stationnement', 'parking interieur': 'Stationnement', 'parking exterieur': 'Stationnement',
  'place de parking': 'Stationnement', 'emplacement': 'Stationnement',
  'archives': 'Archives', 'cave': 'Archives', 'local technique': 'Archives', 'reserve': 'Archives',
  'commerce': 'Commerce', 'boutique': 'Commerce', 'retail': 'Commerce',
  'rie': 'RIE', 'restaurant': 'RIE', 'cafeteria': 'RIE',
}
function normCat(cat) {
  if (!cat) return 'Bureaux'
  const key = String(cat).toLowerCase().trim()
  return CAT_MAP[key] || (Object.keys(CAT_MAP).find(k => key.includes(k)) ? CAT_MAP[Object.keys(CAT_MAP).find(k => key.includes(k))] : cat)
}
function normalizeSurfaces(rows) {
  if (!Array.isArray(rows)) return rows
  return rows.map(r => ({ ...r, categorie: normCat(r.categorie || r.typologie) }))
}

// Deduplicate surfaces_apres : remove rows whose surface_m2 matches sum of avant+delta
function deduplicateSurfacesApres(avant, delta, apres) {
  if (!Array.isArray(apres) || !apres.length) return apres
  if (!Array.isArray(avant) && !Array.isArray(delta)) return apres
  // Build set of m2 values present in avant and delta
  const knownM2 = new Set()
  ;(avant || []).forEach(r => r.surface_m2 && knownM2.add(String(r.surface_m2).trim()))
  ;(delta || []).forEach(r => r.surface_m2 && knownM2.add(String(r.surface_m2).trim()))
  // Remove rows from apres whose m2 is NOT in knownM2 AND matches a computed sum
  // Strategy: if apres has more rows than avant+delta combined, remove rows with m2 that looks like a subtotal
  const avantM2Set = new Set((avant || []).map(r => String(r.surface_m2 || '').trim()))
  const deltaM2Set = new Set((delta || []).map(r => String(r.surface_m2 || '').trim()))
  return apres.filter(r => {
    const m2 = String(r.surface_m2 || '').trim()
    // Keep if m2 exists in avant or delta
    if (avantM2Set.has(m2) || deltaM2Set.has(m2)) return true
    // Remove if m2 is a computed sum of other values (appears nowhere else and is larger than any single entry)
    const num = parseFloat(m2.replace(',', '.'))
    const allNums = [...avantM2Set, ...deltaM2Set].map(v => parseFloat(v.replace(',', '.'))).filter(n => !isNaN(n))
    const maxSingle = Math.max(...allNums, 0)
    if (!isNaN(num) && num > maxSingle) return false // likely a subtotal
    return true
  })
}

// Parse dd/mm/yyyy → Date object
function parseFR(s) {
  if (!s) return null
  const m = String(s).match(/^(\d{2})\/(\d{2})\/(\d{4})$/)
  if (!m) return null
  return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]))
}
// Format Date → dd/mm/yyyy
function fmtFR(d) {
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`
}
// Add N years to a date, keeping same day/month-1 (last day of previous month = expiry convention)
function addYearsExpiry(d, n) {
  // bail starting 01/09/2025 + 3 years → expiry 31/08/2028 (day before same date)
  const result = new Date(d.getFullYear() + n, d.getMonth(), d.getDate() - 1)
  return result
}

function computeBreaks(date_effet_str, date_fin_str, conditions_break_str, existing) {
  const effet = parseFR(date_effet_str)
  const fin   = parseFR(date_fin_str)
  if (!effet || !fin) return existing || []

  const clauseText = (conditions_break_str || '').toLowerCase()

  const candidates = new Set()

  // Detect explicit waiver of triennales ("renonce à sa faculté de résiliation triennale")
  const hasWaiver = /renonce.{0,80}triennale|pas.{0,20}triennale|supprim.{0,20}triennale|faculté.{0,10}résiliation.{0,10}triennale/i.test(clauseText)

  // Detect explicit single break dates ("ne pourra donner congé que pour le jj/mm/aaaa")
  const explicitDates = clauseText.match(/(\d{2})\s*(janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre)\s*(\d{4})/gi) || []
  const monthMap = { janvier:1,février:2,mars:3,avril:4,mai:5,juin:6,juillet:7,août:8,septembre:9,octobre:10,novembre:11,décembre:12 }

  if (!hasWaiver) {
    // Detect triennales (only if no waiver)
    const hasTriennale = /triennale|p.riode.{0,10}3\s*ans/i.test(clauseText) ||
                         /chaque.{0,20}(p.riode|terme|fin)/i.test(clauseText)
    if (hasTriennale) {
      for (let y = 3; y < 9; y += 3) {
        const d = addYearsExpiry(effet, y)
        if (d < fin) candidates.add(fmtFR(d))
      }
    }

    // Detect Nème année patterns
    const yearPattern = /(\d+)[eè][mr]?[eè]?\s+ann[eé]e/gi
    let match
    while ((match = yearPattern.exec(clauseText)) !== null) {
      const n = parseInt(match[1])
      if (n > 0 && n < 9) {
        const d = addYearsExpiry(effet, n)
        if (d < fin) candidates.add(fmtFR(d))
      }
    }
  }

  // Always detect explicit "ne pourra donner congé que pour le DD mois YYYY" patterns
  const explicitPattern = /(\d{1,2})\s+(janvier|f[eé]vrier|mars|avril|mai|juin|juillet|ao[uû]t|septembre|octobre|novembre|d[eé]cembre)\s+(\d{4})/gi
  let exMatch
  while ((exMatch = explicitPattern.exec(clauseText)) !== null) {
    const day = parseInt(exMatch[1])
    const month = monthMap[exMatch[2].toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')] ||
                  monthMap[exMatch[2].toLowerCase()]
    const year = parseInt(exMatch[3])
    if (month && year > 2000) {
      const d = new Date(year, month - 1, day)
      if (d > effet && d < fin) candidates.add(fmtFR(d))
    }
  }

  // Also detect dd/mm/yyyy patterns in clause text
  const datePattern = /(\d{2})\/(\d{2})\/(\d{4})/g
  let dMatch
  while ((dMatch = datePattern.exec(clauseText)) !== null) {
    const d = parseFR(`${dMatch[1]}/${dMatch[2]}/${dMatch[3]}`)
    if (d && d > effet && d < fin) candidates.add(fmtFR(d))
  }

  // If nothing computed and no waiver, fall back to Claude's extracted breaks
  if (!candidates.size && !hasWaiver) return existing || []

  return [...candidates].sort((a, b) => {
    const da = parseFR(a), db = parseFR(b)
    if (!da || !db) return 0
    return da - db
  })
}

// Remove parking rows that have no useful data (no surface_m2 and no loyer_annuel)
function cleanSurfaces(rows) {
  if (!Array.isArray(rows)) return rows
  return rows.filter(r => {
    const cat = (r.categorie || r.typologie || '').toLowerCase()
    const isPark = cat.includes('station') || cat.includes('parking') || cat.includes('place')
    if (!isPark) return true
    // Keep parking row only if it has a surface or a loyer
    return (r.surface_m2 && String(r.surface_m2).trim() !== '' && String(r.surface_m2) !== '0') ||
           (r.loyer_annuel && String(r.loyer_annuel).trim() !== '')
  })
}
function mergeSurfacesByCategory(rows) {
  if (!Array.isArray(rows) || rows.length <= 1) return rows
  const map = new Map()
  rows.forEach(r => {
    const cat = r.categorie || r.typologie || '—'
    if (!map.has(cat)) {
      map.set(cat, { ...r })
    } else {
      const existing = map.get(cat)
      const a = parseFloat(String(existing.surface_m2 || '').replace(',', '.')) || 0
      const b = parseFloat(String(r.surface_m2 || '').replace(',', '.')) || 0
      if (a > 0 && b > 0) {
        existing.surface_m2 = String(Math.round((a + b) * 100) / 100)
        existing.niveau = null // mixed
        // Sum loyer_annuel too
        const la = parseAmount(existing.loyer_annuel), lb = parseAmount(r.loyer_annuel)
        if (la !== null && lb !== null) existing.loyer_annuel = String(la + lb)
      }
    }
  })
  return [...map.values()]
}


// Compute loyer unitaire if missing from surfaces_detail
function computeUnitPrices(surfaces, parkingNbPlaces, parkingLoyer) {
  if (!Array.isArray(surfaces)) return surfaces
  return surfaces.map(row => {
    if (row.prix_unitaire) return row
    const loyer = parseAmount(row.loyer_annuel)
    const surf  = parseFloat(String(row.surface_m2 || '').replace(',', '.')) || 0
    const isPark = (row.categorie || '').toLowerCase().includes('station')
    // For parking rows without surface: try to derive from total places
    if (isPark && !surf && loyer !== null) {
      // Count nb places from parking_nb_places global field if only one parking row
      return { ...row }
    }
    if (loyer !== null && surf > 0) {
      return { ...row, prix_unitaire: String(Math.round(loyer / surf)) }
    }
    return row
  })
}

// Compute parking unit price: loyer / nb_places
function computeParkingUnitPrice(parkingStr, surfaces) {
  // Extract total loyer from parking rows in surfaces_detail
  const parkRows = Array.isArray(surfaces) ? surfaces.filter(r => (r.categorie || '').toLowerCase().includes('station')) : []
  const totalLoyerPark = parkRows.reduce((acc, r) => acc + (parseAmount(r.loyer_annuel) || 0), 0)
  // Extract nb places from string like "291 places" or "291"
  const nbMatch = String(parkingStr || '').match(/(\d+)/)
  const nb = nbMatch ? parseInt(nbMatch[1]) : 0
  if (nb > 0 && totalLoyerPark > 0) {
    return Math.round(totalLoyerPark / nb)
  }
  return null
}

// INSEE BDM series IDs for real estate indices
const INSEE_SERIES = {
  ILAT: '001769193',
  ILC:  '001629154',
  ICC:  '001771006',
  IRL:  '001671963',
  IPC:  '001759970',
}

// Convert dd/mm/yyyy to quarter string like "2025-Q3"
function dateToQuarter(dateStr) {
  const m = String(dateStr || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/)
  if (!m) return null
  const month = parseInt(m[2])
  const year  = parseInt(m[3])
  const q = Math.ceil(month / 3)
  return `${year}-Q${q}`
}

// Static INSEE index table (last known values, updated periodically)
// Source: INSEE BDM - indices immobiliers tertiaires
const INSEE_STATIC = {
  ILAT: [
    { q: '1T2015', v: 107.44 }, { q: '2T2015', v: 107.53 }, { q: '3T2015', v: 107.43 }, { q: '4T2015', v: 107.34 },
    { q: '1T2016', v: 107.22 }, { q: '2T2016', v: 107.28 }, { q: '3T2016', v: 107.50 }, { q: '4T2016', v: 107.64 },
    { q: '1T2017', v: 107.87 }, { q: '2T2017', v: 108.28 }, { q: '3T2017', v: 108.65 }, { q: '4T2017', v: 109.06 },
    { q: '1T2018', v: 109.57 }, { q: '2T2018', v: 110.25 }, { q: '3T2018', v: 110.82 }, { q: '4T2018', v: 111.19 },
    { q: '1T2019', v: 111.59 }, { q: '2T2019', v: 112.07 }, { q: '3T2019', v: 112.47 }, { q: '4T2019', v: 112.78 },
    { q: '1T2020', v: 113.08 }, { q: '2T2020', v: 112.83 }, { q: '3T2020', v: 112.91 }, { q: '4T2020', v: 113.00 },
    { q: '1T2021', v: 113.18 }, { q: '2T2021', v: 113.65 }, { q: '3T2021', v: 114.49 }, { q: '4T2021', v: 115.59 },
    { q: '1T2022', v: 117.22 }, { q: '2T2022', v: 119.76 }, { q: '3T2022', v: 122.42 }, { q: '4T2022', v: 124.87 },
    { q: '1T2023', v: 127.22 }, { q: '2T2023', v: 129.02 }, { q: '3T2023', v: 130.44 }, { q: '4T2023', v: 131.34 },
    { q: '1T2024', v: 132.28 }, { q: '2T2024', v: 133.01 }, { q: '3T2024', v: 133.56 }, { q: '4T2024', v: 133.99 },
    { q: '1T2025', v: 134.48 }, { q: '2T2025', v: 135.02 },
  ],
  ILC: [
    { q: '1T2015', v: 108.82 }, { q: '2T2015', v: 109.04 }, { q: '3T2015', v: 108.97 }, { q: '4T2015', v: 108.82 },
    { q: '1T2016', v: 108.65 }, { q: '2T2016', v: 108.89 }, { q: '3T2016', v: 109.14 }, { q: '4T2016', v: 109.35 },
    { q: '1T2017', v: 109.70 }, { q: '2T2017', v: 110.14 }, { q: '3T2017', v: 110.47 }, { q: '4T2017', v: 110.82 },
    { q: '1T2018', v: 111.38 }, { q: '2T2018', v: 111.95 }, { q: '3T2018', v: 112.33 }, { q: '4T2018', v: 112.65 },
    { q: '1T2019', v: 113.07 }, { q: '2T2019', v: 113.40 }, { q: '3T2019', v: 113.65 }, { q: '4T2019', v: 113.98 },
    { q: '1T2020', v: 114.15 }, { q: '2T2020', v: 113.73 }, { q: '3T2020', v: 113.82 }, { q: '4T2020', v: 113.94 },
    { q: '1T2021', v: 114.20 }, { q: '2T2021', v: 114.73 }, { q: '3T2021', v: 115.62 }, { q: '4T2021', v: 116.75 },
    { q: '1T2022', v: 118.45 }, { q: '2T2022', v: 121.04 }, { q: '3T2022', v: 123.69 }, { q: '4T2022', v: 126.12 },
    { q: '1T2023', v: 128.95 }, { q: '2T2023', v: 130.21 }, { q: '3T2023', v: 131.44 }, { q: '4T2023', v: 132.03 },
    { q: '1T2024', v: 132.97 }, { q: '2T2024', v: 133.82 }, { q: '3T2024', v: 134.51 }, { q: '4T2024', v: 135.12 },
    { q: '1T2025', v: 135.89 }, { q: '2T2025', v: 136.74 },
  ],
  ICC: [
    { q: '1T2015', v: 1640 }, { q: '2T2015', v: 1640 }, { q: '3T2015', v: 1637 }, { q: '4T2015', v: 1635 },
    { q: '1T2016', v: 1634 }, { q: '2T2016', v: 1638 }, { q: '3T2016', v: 1641 }, { q: '4T2016', v: 1644 },
    { q: '1T2017', v: 1649 }, { q: '2T2017', v: 1658 }, { q: '3T2017', v: 1665 }, { q: '4T2017', v: 1672 },
    { q: '1T2018', v: 1683 }, { q: '2T2018', v: 1697 }, { q: '3T2018', v: 1707 }, { q: '4T2018', v: 1714 },
    { q: '1T2019', v: 1720 }, { q: '2T2019', v: 1730 }, { q: '3T2019', v: 1737 }, { q: '4T2019', v: 1742 },
    { q: '1T2020', v: 1748 }, { q: '2T2020', v: 1741 }, { q: '3T2020', v: 1749 }, { q: '4T2020', v: 1757 },
    { q: '1T2021', v: 1770 }, { q: '2T2021', v: 1797 }, { q: '3T2021', v: 1835 }, { q: '4T2021', v: 1888 },
    { q: '1T2022', v: 1958 }, { q: '2T2022', v: 2023 }, { q: '3T2022', v: 2059 }, { q: '4T2022', v: 2075 },
    { q: '1T2023', v: 2053 }, { q: '2T2023', v: 2071 }, { q: '3T2023', v: 2089 }, { q: '4T2023', v: 2095 },
    { q: '1T2024', v: 2107 }, { q: '2T2024', v: 2119 }, { q: '3T2024', v: 2128 }, { q: '4T2024', v: 2134 },
    { q: '1T2025', v: 2141 },
  ],
  IRL: [
    { q: '1T2015', v: 125.01 }, { q: '2T2015', v: 125.25 }, { q: '3T2015', v: 125.31 }, { q: '4T2015', v: 125.29 },
    { q: '1T2016', v: 125.24 }, { q: '2T2016', v: 125.19 }, { q: '3T2016', v: 125.29 }, { q: '4T2016', v: 125.40 },
    { q: '1T2017', v: 125.55 }, { q: '2T2017', v: 125.90 }, { q: '3T2017', v: 126.25 }, { q: '4T2017', v: 126.50 },
    { q: '1T2018', v: 126.83 }, { q: '2T2018', v: 127.22 }, { q: '3T2018', v: 127.60 }, { q: '4T2018', v: 128.02 },
    { q: '1T2019', v: 128.45 }, { q: '2T2019', v: 128.93 }, { q: '3T2019', v: 129.38 }, { q: '4T2019', v: 129.72 },
    { q: '1T2020', v: 130.10 }, { q: '2T2020', v: 130.26 }, { q: '3T2020', v: 130.57 }, { q: '4T2020', v: 130.88 },
    { q: '1T2021', v: 131.12 }, { q: '2T2021', v: 131.67 }, { q: '3T2021', v: 132.43 }, { q: '4T2021', v: 133.93 },
    { q: '1T2022', v: 135.84 }, { q: '2T2022', v: 138.50 }, { q: '3T2022', v: 140.59 }, { q: '4T2022', v: 142.11 },
    { q: '1T2023', v: 143.51 }, { q: '2T2023', v: 145.19 }, { q: '3T2023', v: 146.52 }, { q: '4T2023', v: 147.11 },
    { q: '1T2024', v: 147.65 }, { q: '2T2024', v: 148.43 }, { q: '3T2024', v: 149.02 }, { q: '4T2024', v: 149.61 },
    { q: '1T2025', v: 150.14 },
  ],
}

// Convert dd/mm/yyyy to "NT" quarter string like "3T2025"
function dateToQuarterLabel(dateStr) {
  const m = String(dateStr || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/)
  if (!m) return null
  const month = parseInt(m[2])
  const year  = parseInt(m[3])
  const q = Math.ceil(month / 3)
  return `${q}T${year}`
}

// Find last published index at or before the given date from static table
// Tries INSEE BDM API first, falls back to static table
async function fetchInseeIndex(indice, dateStr) {
  const key = indice?.toUpperCase()
  if (!key) return null
  const targetLabel = dateToQuarterLabel(dateStr)
  if (!targetLabel) return null

  // Try INSEE BDM API (works if CORS allows)
  const series = INSEE_SERIES[key]
  if (series) {
    try {
      const m = targetLabel.match(/(\d)T(\d{4})/)
      if (m) {
        const startYear = parseInt(m[2]) - 1
        const endPeriod = `${m[2]}-Q${m[1]}`
        const url = `https://api.insee.fr/series/BDM/V1/data/SERIES_BDM/${series}?startPeriod=${startYear}-Q1&endPeriod=${endPeriod}`
        const res = await fetch(url, { headers: { Accept: 'application/json' }, signal: AbortSignal.timeout(4000) })
        if (res.ok) {
          const data = await res.json()
          const obs = data?.GenericData?.DataSet?.Series?.Obs
          if (Array.isArray(obs) && obs.length) {
            const last = obs[obs.length - 1]
            const period = last?.['@TIME_PERIOD'] || ''
            const value  = last?.ObsValue?.['@value'] || last?.['@OBS_VALUE']
            const pMatch = String(period).match(/(\d{4})-Q(\d)/)
            const label  = pMatch ? `${pMatch[2]}T${pMatch[1]}` : period
            if (value) return { value: parseFloat(value), label, source: 'INSEE' }
          }
        }
      }
    } catch { /* fall through to static */ }
  }

  // Fallback: static table
  const table = INSEE_STATIC[key]
  if (!table) return null
  // Parse target: "3T2025" → year=2025, q=3
  const tm = targetLabel.match(/(\d)T(\d{4})/)
  if (!tm) return null
  const targetYear = parseInt(tm[2]), targetQ = parseInt(tm[1])
  // Find last entry <= target
  const candidates = table.filter(row => {
    const rm = row.q.match(/(\d)T(\d{4})/)
    if (!rm) return false
    const y = parseInt(rm[2]), q = parseInt(rm[1])
    return y < targetYear || (y === targetYear && q <= targetQ)
  })
  if (!candidates.length) return null
  const last = candidates[candidates.length - 1]
  return { value: last.v, label: last.q, source: 'INSEE (table)' }
}

function sanitizeExtracted(data) {
  if (!data || typeof data !== 'object') return data
  const d = { ...data }
  d.break_options = ensureArray(d.break_options)
  // Enrichir les breaks par calcul côté code — fiable à 100%
  if (d.date_effet || d.date_fin) {
    d.break_options = computeBreaks(d.date_effet, d.date_fin, d.conditions_break, d.break_options)
  }
  const cs = rows => cleanSurfaces(normalizeSurfaces(rows))
  d.surfaces_detail    = cs(ensureArray(d.surfaces_detail))
  d.franchise_periodes = ensureArray(d.franchise_periodes)
  d.indemnites         = ensureArray(d.indemnites)
  d.surfaces_delta          = cs(ensureArray(d.surfaces_delta))
  d.participations_travaux  = ensureArray(d.participations_travaux)
  d.paliers_loyer           = ensureArray(d.paliers_loyer)
  d.abattements             = ensureArray(d.abattements)
  d.indemnites_break        = ensureArray(d.indemnites_break)
  d.indemnites_restitution  = ensureArray(d.indemnites_restitution)
  if (d.champs_modifies) {
    d.champs_modifies.participations_travaux = ensureArray(d.champs_modifies?.participations_travaux)
    d.champs_modifies.paliers_loyer          = ensureArray(d.champs_modifies?.paliers_loyer)
    d.champs_modifies.abattements            = ensureArray(d.champs_modifies?.abattements)
    d.champs_modifies.indemnites_break       = ensureArray(d.champs_modifies?.indemnites_break)
  }
  d.surfaces_avant  = cs(ensureArray(d.surfaces_avant))
  d.surfaces_apres  = cs(mergeSurfacesByCategory(ensureArray(d.surfaces_apres)))
  if (d.champs_modifies) {
    d.champs_modifies = { ...d.champs_modifies }
    d.champs_modifies.break_options      = ensureArray(d.champs_modifies.break_options)
    d.champs_modifies.surfaces_detail    = cs(ensureArray(d.champs_modifies.surfaces_detail))
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

// Raw single call — used internally and for detect
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

// Dual-pass extraction for bail: structural + financial in parallel, then merge

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

// Parse a verbose text into bullet items splitting on ; or \n or " - "
function parseBullets(text) {
  if (!text) return []
  const s = (safeStr(text) || '').trim()
  // 1. Split on explicit semicolons
  if (s.includes(';')) {
    return s.split(/\s*;\s*/).map(p => p.trim()).filter(p => p.length > 2)
  }
  // 2. Split on newlines
  if (s.includes('\n')) {
    return s.split(/\n+/).map(p => p.trim()).filter(p => p.length > 2)
  }
  // 3. Split on ". " followed by capital letter or "À" (French sentence boundary)
  // but keep abbreviations like "art. 606", "HT.", "CC." intact
  const abbrevPattern = /\b(art|al|n°|HT|HC|TTC|CC|CGI|m²|cf|etc|ex|vs|p|pp|vol|réf)\b\.?\s*$/i
  const parts = []
  let current = ''
  // Walk char by char detecting sentence ends
  const sentences = s.split(/(?<=[\w\)\]°%€])\.\s+(?=[A-ZÀÂÄÉÈÊËÎÏÔÙÛÜ"«])/u)
  if (sentences.length > 1) {
    return sentences.map(p => p.trim()).filter(p => p.length > 2)
  }
  // 4. Fallback: single block
  return [s]
}

function BulletField({ label, value, full }) {
  const s = safeStr(value)
  if (!s) return null
  const items = parseBullets(s)
  return (
    <div className={`field verbose-field${full ? ' full' : ''}`}>
      <div className="field-lbl" style={{ marginBottom: '8px' }}>{label}</div>
      {items.length > 1 ? (
        <ul style={{ margin: 0, padding: '0 0 0 14px', listStyle: 'none' }}>
          {items.map((item, i) => (
            <li key={i} style={{ display: 'flex', gap: '8px', fontSize: '12.5px', color: 'var(--text2)', lineHeight: 1.6, marginBottom: i < items.length - 1 ? '4px' : 0 }}>
              <span style={{ color: 'var(--border2)', flexShrink: 0, fontWeight: 700, marginTop: '1px' }}>—</span>
              <span>{item}</span>
            </li>
          ))}
        </ul>
      ) : (
        <div className="field-val verbose">{s}</div>
      )}
    </div>
  )
}

function Field({ label, value, mono, verbose }) {
  const safe = safeStr(value)
  return (
    <div className="field">
      <div className="field-lbl">{label}</div>
      <div className={`field-val${!safe ? ' empty' : mono ? ' mono' : verbose ? ' verbose' : ''}`}>{safe || 'Non renseigné'}</div>
    </div>
  )
}

function PairBlock({ keyLabel, keyValue, keyMono, verboseLabel, verboseValue }) {
  const safeKey = safeStr(keyValue)
  const safeVerbose = safeStr(verboseValue)
  return (
    <div className="pair-block full">
      <div className="pair-key">
        <div className="field-lbl">{keyLabel}</div>
        <div className={`field-val${!safeKey ? ' empty' : keyMono ? ' mono' : ''}`}>{safeKey || 'Non renseigné'}</div>
      </div>
      <div className="pair-verbose">
        <div className="field-lbl">{verboseLabel}</div>
        <div className={`field-val${!safeVerbose ? ' empty' : ' verbose'}`}>{safeVerbose || 'Non renseigné'}</div>
      </div>
    </div>
  )
}

function SurfaceTable({ surfaces }) {
  const safe = Array.isArray(surfaces) ? surfaces : []
  if (!safe.length) return null
  const isPark = r => { const cat = (r.categorie || r.typologie || '').toLowerCase(); return cat.includes('station') || cat.includes('parking') || cat.includes('place') }
  const mainRows = safe.filter(r => !isPark(r))
  const parkRows = safe.filter(r => isPark(r))
  const total = mainRows.reduce((acc, r) => acc + (parseFloat(String(r.surface_m2 || '').replace(/[^0-9.]/g, '')) || 0), 0)
  const totalLoyer = safe.reduce((acc, r) => acc + (parseAmount(r.loyer_annuel) || 0), 0)
  const parkTotalLoyer = parkRows.reduce((acc, r) => acc + (parseAmount(r.loyer_annuel) || 0), 0)

  // Compute unit price for a row if missing
  const unitPrice = r => {
    if (r.prix_unitaire) return parseAmount(r.prix_unitaire)
    const loyer = parseAmount(r.loyer_annuel)
    const surf  = parseFloat(String(r.surface_m2 || '').replace(',', '.')) || 0
    if (loyer !== null && surf > 0) return Math.round(loyer / surf)
    return null
  }

  return (
    <div>
      {mainRows.length > 0 && (
        <div className="table-wrap">
          <table className="indemnites-table">
            <thead>
              <tr>
                <th>Catégorie</th><th>Niveau / Localisation</th>
                <th style={{ textAlign: 'right' }}>Surface (m²)</th>
                <th style={{ textAlign: 'right' }}>Prix (€/m²/an)</th>
                <th style={{ textAlign: 'right' }}>Loyer annuel (€)</th>
              </tr>
            </thead>
            <tbody>
              {mainRows.map((row, i) => {
                const up = unitPrice(row)
                return (
                  <tr key={i}>
                    <td style={{ fontWeight: 500 }}>{row.categorie || row.typologie || '—'}</td>
                    <td style={{ color: 'var(--text2)' }}>{row.niveau || row.localisation || '—'}</td>
                    <td style={{ textAlign: 'right', fontWeight: 500 }}>{row.surface_m2 ? `${row.surface_m2} m²` : '—'}</td>
                    <td style={{ textAlign: 'right', color: row.prix_unitaire ? 'var(--text)' : 'var(--text3)', fontStyle: row.prix_unitaire ? 'normal' : 'italic' }}>
                      {up ? `${up.toLocaleString('fr-FR')} €` : '—'}
                      {!row.prix_unitaire && up && <span title="Calculé" style={{ fontSize: '10px', marginLeft: '3px' }}>*</span>}
                    </td>
                    <td style={{ textAlign: 'right', fontWeight: 500 }}>{row.loyer_annuel ? fmtEur(row.loyer_annuel) : '—'}</td>
                  </tr>
                )
              })}
            </tbody>
            {total > 0 && (
              <tfoot>
                <tr style={{ borderTop: '1px solid var(--border2)' }}>
                  <td colSpan={2} style={{ fontWeight: 600, padding: '8px 10px' }}>Total bureaux / locaux</td>
                  <td style={{ textAlign: 'right', fontWeight: 600, padding: '8px 10px' }}>{total.toLocaleString('fr-FR')} m²</td>
                  <td />
                  <td style={{ textAlign: 'right', fontWeight: 600, padding: '8px 10px' }}>
                    {mainRows.reduce((a,r) => a + (parseAmount(r.loyer_annuel)||0), 0) > 0
                      ? fmtEur(mainRows.reduce((a,r) => a + (parseAmount(r.loyer_annuel)||0), 0)) : '—'}
                  </td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      )}
      {parkRows.length > 0 && (
        <div style={{ marginTop: mainRows.length > 0 ? '10px' : 0 }}>
          <table className="indemnites-table">
            <thead>
              <tr>
                <th>Stationnement</th><th>Localisation</th>
                <th style={{ textAlign: 'right' }}>Nb places</th>
                <th style={{ textAlign: 'right' }}>Prix (€/place/an)</th>
                <th style={{ textAlign: 'right' }}>Loyer annuel (€)</th>
              </tr>
            </thead>
            <tbody>
              {parkRows.map((row, i) => {
                const loyer = parseAmount(row.loyer_annuel)
                const surf  = parseFloat(String(row.surface_m2 || '').replace(/[^0-9.]/g,'')) || 0
                const up    = (loyer !== null && surf > 0) ? Math.round(loyer / surf) : null
                return (
                  <tr key={i}>
                    <td style={{ fontWeight: 500 }}>{row.categorie || 'Stationnement'}</td>
                    <td style={{ color: 'var(--text2)' }}>{row.niveau || row.localisation || '—'}</td>
                    <td style={{ textAlign: 'right' }}>{row.surface_m2 ? `${row.surface_m2} pl.` : '—'}</td>
                    <td style={{ textAlign: 'right', color: up ? 'var(--text)' : 'var(--text3)', fontStyle: up ? 'normal' : 'italic' }}>
                      {up ? `${up.toLocaleString('fr-FR')} €` : '—'}
                      {up && <span title="Calculé" style={{ fontSize: '10px', marginLeft: '3px' }}>*</span>}
                    </td>
                    <td style={{ textAlign: 'right', fontWeight: 500 }}>{loyer !== null ? fmtEur(loyer) : '—'}</td>
                  </tr>
                )
              })}
            </tbody>
            {parkTotalLoyer > 0 && (
              <tfoot>
                <tr style={{ borderTop: '1px solid var(--border2)' }}>
                  <td colSpan={4} style={{ fontWeight: 600, padding: '8px 10px' }}>Total stationnement</td>
                  <td style={{ textAlign: 'right', fontWeight: 600, padding: '8px 10px' }}>{fmtEur(parkTotalLoyer)}</td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      )}
      {totalLoyer > 0 && (mainRows.length > 0 && parkRows.length > 0) && (
        <div style={{ textAlign: 'right', padding: '6px 10px', fontWeight: 700, fontSize: '13px', borderTop: '2px solid var(--border2)', marginTop: '2px' }}>
          Total général : {fmtEur(totalLoyer)}
        </div>
      )}
      {(mainRows.length > 0 || parkRows.length > 0) && (
        <div style={{ fontSize: '10px', color: 'var(--text3)', marginTop: '4px' }}>* Prix unitaire calculé (loyer annuel ÷ surface/places)</div>
      )}
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
            <th>Surface assiette</th>
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
              <td style={{ color: 'var(--text2)' }}>{row.surface_assiette || '—'}</td>
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
              <td style={{ color: 'var(--text2)', fontStyle: row.condition ? 'normal' : 'italic' }}>{safeStr(row.condition) || '—'}</td>
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
              <td>{safeStr(row.motif) || '—'}</td>
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
              <td style={{ color: 'var(--text2)' }}>{safeStr(row.date_limite) || '—'}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}


// Extract short name: stop at comma, parenthesis, or legal form keywords
function shortPartyName(s) {
  if (!s) return s
  return s
    .split(/,|\(|(?<=\S)\s+(SAS|SA|SARL|SCI|SASU|SNC|GIE|EURL|SE|société|Société|SOCIÉTÉ|S\.A\.|S\.A\.S\.|S\.C\.I\.)/)[0]
    .trim()
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

  const [inseeIndex, setInseeIndex] = useState(null)
  const [inseeLoading, setInseeLoading] = useState(false)
  useEffect(() => {
    setInseeIndex(null)
    const indice = d.indexation_indice
    if (!indice || !INSEE_SERIES[indice]) return
    if (d.indexation_valeur_base) {
      setInseeIndex({ value: parseFloat(d.indexation_valeur_base), label: d.indexation_trimestre_base || '', source: 'bail' })
      return
    }
    setInseeLoading(true)
    fetchInseeIndex(indice, d.date_signature || d.date_effet).then(res => {
      setInseeIndex(res)
      setInseeLoading(false)
    })
  }, [item.id, d.indexation_indice, d.indexation_valeur_base, d.date_signature])

  // Enrichir les breaks à l'affichage aussi (données déjà en base non recalculées)
  const breaks = computeBreaks(d.date_effet, d.date_fin, d.conditions_break, d.break_options || [])

  // Clean surfaces at display time too (for data already in DB)
  const cs = rows => cleanSurfaces(normalizeSurfaces(Array.isArray(rows) ? rows : []))
  d.surfaces_detail = cs(d.surfaces_detail)
  if (d.surfaces_avant) d.surfaces_avant = cs(d.surfaces_avant)
  if (d.surfaces_delta) d.surfaces_delta = cs(d.surfaces_delta)
  if (d.surfaces_apres) d.surfaces_apres = cs(mergeSurfacesByCategory(d.surfaces_apres))

  const indemnites = d.indemnites.length > 0 ? d.indemnites : null

  const pills = []
  if (d.indexation) {
    const idx = safeStr(d.indexation)?.toLowerCase() || ''
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

  const primaryDates = [
    d.date_effet ? { key: 'date_effet', label: "Prise d'effet", type: 'primary' } : null,
    ...breaks.map((br, i) => ({ key: `break_${i}`, label: `Break option${breaks.length > 1 ? ' '+( i+1) : ''}`, val: br, type: 'break' })),
    d.date_fin   ? { key: 'date_fin',   label: 'Expiration',    type: 'primary' } : null,
  ].filter(Boolean)

  const secondaryDates = [
    { key: 'date_signature',      label: 'Signature' },
    { key: 'notice',              label: 'Préavis' },
    { key: 'date_limite_travaux', label: 'Date limite travaux preneur' },
  ].filter(f => d[f.key])

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
          {meta.date_effet_avenant && <span className="av-banner-date">Effet : {normalizeDate(meta.date_effet_avenant)}</span>}
        </div>
      )}

      {/* Parties */}
      {(show('preneur') || show('bailleur')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Parties</div></div>
          <div className="gx">
            {show('preneur') && <div className="party-card"><div className="party-role">Preneur</div><div className="party-name">{shortPartyName(d.preneur) || <span style={{ color: 'var(--text3)', fontStyle: 'italic', fontWeight: 400 }}>Non renseigné</span>}</div></div>}
            {show('bailleur') && <div className="party-card"><div className="party-role">Bailleur</div><div className="party-name">{shortPartyName(d.bailleur) || <span style={{ color: 'var(--text3)', fontStyle: 'italic', fontWeight: 400 }}>Non renseigné</span>}</div></div>}
            {show('garant') && d.garant && <div className="party-card" style={{gridColumn:'1/-1'}}><div className="party-role">Garant / Caution</div><div className="party-name">{shortPartyName(d.garant)}</div></div>}
          </div>
        </div>
      )}

      {/* Contrat */}
      {(show('type_bail') || show('duree_totale')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Contrat et durée</div></div>
          <div className="gx">
            <Field label="Type de contrat" value={d.type_bail} />
            <Field label="Durée totale" value={d.duree_totale} />
            <Field label="Durée ferme" value={d.duree_ferme} />
          </div>
        </div>
      )}

      {/* Dates */}
      {(primaryDates.length > 0 || secondaryDates.length > 0 || show('conditions_break')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Dates clés</div></div>
          {/* Rangée principale : effet — breaks — expiration */}
          {primaryDates.length > 0 && (
            <div className="date-strip" style={{ gridTemplateColumns: `repeat(${Math.min(primaryDates.length, 5)}, 1fr)`, marginBottom: '12px' }}>
              {primaryDates.map(f => (
                <div key={f.key} className={`date-card${f.type === 'break' ? ' date-card-break' : ''}`}>
                  <div className="date-lbl">
                    {f.type === 'break' && <span className="break-tag">B{breaks.length > 1 ? f.key.split('_')[1]*1+1 : ''}</span>}
                    {' '}{f.label}
                  </div>
                  <div className={`date-val${f.type === 'break' ? ' break' : ''}`}>{f.val || d[f.key]}</div>
                  {f.type === 'break' && d.notice && <div style={{ fontSize: '11px', color: 'var(--text3)', marginTop: '4px' }}>Préavis : {d.notice}</div>}
                </div>
              ))}
            </div>
          )}
          {/* Rangée secondaire : signature, congé, travaux */}
          {secondaryDates.length > 0 && (
            <div className="gx" style={{ marginBottom: '8px' }}>
              {secondaryDates.map(f => (
                <div key={f.key} style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 'var(--r)', padding: '8px 14px', minWidth: '160px' }}>
                  <div className="field-lbl">{f.label}</div>
                  <div className="field-val">{d[f.key]}</div>
                </div>
              ))}
            </div>
          )}
          {show('conditions_break') && d.conditions_break && (
            <div style={{ marginTop: '8px' }}>
              <BulletField label="Détail échéances" value={d.conditions_break} full />
            </div>
          )}
        </div>
      )}

      {/* Mise à disposition anticipée */}
      {d.mise_a_disposition && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Mise à disposition anticipée</div></div>
          <div className="gx">
            {d.mise_a_disposition.date_debut && <Field label="Date de début" value={d.mise_a_disposition.date_debut} />}
            {d.mise_a_disposition.date_fin   && <Field label="Date de fin"   value={d.mise_a_disposition.date_fin} />}
            {d.mise_a_disposition.loyer_paye && (
              <div className="field">
                <div className="field-lbl">Loyer payé</div>
                <div className="field-val">
                  <span className={`pill ${d.mise_a_disposition.loyer_paye === 'Oui' ? 'pill-danger' : d.mise_a_disposition.loyer_paye === 'Non' ? 'pill-green' : 'pill-blue'}`}>
                    {d.mise_a_disposition.loyer_paye}
                  </span>
                </div>
              </div>
            )}
            {d.mise_a_disposition.charges_payees && (
              <div className="field">
                <div className="field-lbl">Charges payées</div>
                <div className="field-val">
                  <span className={`pill ${d.mise_a_disposition.charges_payees === 'Oui' ? 'pill-danger' : d.mise_a_disposition.charges_payees === 'Non' ? 'pill-green' : 'pill-blue'}`}>
                    {d.mise_a_disposition.charges_payees}
                  </span>
                </div>
              </div>
            )}
          </div>
          {d.mise_a_disposition.conditions && (
            <Field label="Conditions financières" value={safeStr(d.mise_a_disposition.conditions)} verbose full />
          )}
        </div>
      )}


      {(show('surface_totale_m2') || show('parking_nb_places') || show('rie')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Surfaces</div></div>
          <div className="gx">
            {show('surface_totale_m2') && (
              <div className="field">
                <div className="field-lbl">Surface totale louée</div>
                <div className="field-val" style={{ fontSize: '17px', fontWeight: 700, letterSpacing: '-0.02em' }}>{d.surface_totale_m2 ? `${d.surface_totale_m2} m²` : '—'}</div>
              </div>
            )}
            {show('parking_nb_places') && (() => {
              const pkUnit = computeParkingUnitPrice(d.parking_nb_places, d.surfaces_detail)
              return (
                <div className="field">
                  <div className="field-lbl">Stationnement</div>
                  <div className="field-val">
                    {parseParkingShort(d.parking_nb_places) || '—'}
                    {pkUnit && <span style={{ fontSize: '12px', fontWeight: 400, color: 'var(--text2)', marginLeft: '8px' }}>{pkUnit.toLocaleString('fr-FR')} €/place/an</span>}
                  </div>
                </div>
              )
            })()}
            {show('rie') && d.rie && <Field label="RIE" value={d.rie} />}
          </div>
          {/* Bloc modification surfaces (avenants uniquement) */}
          {isAv && (() => {
            const sct = item.data?.surface_change_type
            const avant = item.data?.surfaces_avant
            const delta = item.data?.surfaces_delta
            const apres = item.data?.surfaces_apres
            if (!sct || sct === 'inchangee') return null
            const labelMap = {
              ajout:        { txt: 'Ajout de surfaces',         cls: 'pill-green'  },
              retrait:      { txt: 'Retrait de surfaces',       cls: 'pill-danger' },
              substitution: { txt: 'Substitution de surfaces',  cls: 'pill-blue'   },
              mixte:        { txt: 'Modification mixte',        cls: 'pill-blue'   },
            }
            const lbl = labelMap[sct] || { txt: sct, cls: 'pill-blue' }

            function SurfMiniTable({ rows, accentSens }) {
              if (!Array.isArray(rows) || !rows.length) return <span style={{ fontSize: '12px', color: 'var(--text3)', fontStyle: 'italic' }}>—</span>
              const isPark = r => { const cat = (r.categorie || r.typologie || '').toLowerCase(); return cat.includes('station') || cat.includes('parking') || cat.includes('place') }
              const mainRows = rows.filter(r => !isPark(r))
              const parkRows = rows.filter(r => isPark(r))
              const totalM2 = mainRows.reduce((acc, r) => acc + (parseFloat(String(r.surface_m2 || '').replace(/[^0-9.]/g, '')) || 0), 0)
              const parkCount = parkRows.reduce((acc, r) => acc + (parseFloat(String(r.surface_m2 || '').replace(/[^0-9.]/g, '')) || 0), 0)
              return (
                <div>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                  <thead>
                    <tr style={{ borderBottom: '1px solid var(--border2)' }}>
                      {accentSens && <th style={{ padding: '4px 6px', textAlign: 'left', fontWeight: 600, color: 'var(--text3)', fontSize: '10px', textTransform: 'uppercase', letterSpacing: '.05em' }}>Sens</th>}
                      <th style={{ padding: '4px 6px', textAlign: 'left', fontWeight: 600, color: 'var(--text3)', fontSize: '10px', textTransform: 'uppercase', letterSpacing: '.05em' }}>Catégorie</th>
                      <th style={{ padding: '4px 6px', textAlign: 'left', fontWeight: 600, color: 'var(--text3)', fontSize: '10px', textTransform: 'uppercase', letterSpacing: '.05em' }}>Niveau</th>
                      <th style={{ padding: '4px 6px', textAlign: 'right', fontWeight: 600, color: 'var(--text3)', fontSize: '10px', textTransform: 'uppercase', letterSpacing: '.05em' }}>m²</th>
                      <th style={{ padding: '4px 6px', textAlign: 'right', fontWeight: 600, color: 'var(--text3)', fontSize: '10px', textTransform: 'uppercase', letterSpacing: '.05em' }}>Loyer/an</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((row, i) => (
                      <tr key={i} style={{ borderBottom: '1px solid var(--border)' }}>
                        {accentSens && (
                          <td style={{ padding: '5px 6px' }}>
                            <span style={{ fontSize: '10px', fontWeight: 700, padding: '1px 5px', borderRadius: '3px',
                              background: row.sens === 'retrait' ? 'var(--danger-bg)' : 'var(--success-bg)',
                              color: row.sens === 'retrait' ? 'var(--danger)' : 'var(--success)' }}>
                              {row.sens === 'retrait' ? '−' : '+'}
                            </span>
                          </td>
                        )}
                        <td style={{ padding: '5px 6px', fontWeight: 500 }}>{row.categorie || row.typologie || '—'}</td>
                        <td style={{ padding: '5px 6px', color: 'var(--text2)', whiteSpace: 'nowrap' }}>{safeStr(row.niveau || row.localisation) || '—'}</td>
                        <td style={{ padding: '5px 6px', textAlign: 'right', fontWeight: 600, whiteSpace: 'nowrap' }}>{row.surface_m2 ? `${row.surface_m2} m²` : '—'}</td>
                        <td style={{ padding: '5px 6px', textAlign: 'right' }}>{row.loyer_annuel ? fmtEur(row.loyer_annuel) : '—'}</td>
                      </tr>
                    ))}
                  </tbody>
                  {totalM2 > 0 && (
                    <tfoot>
                      <tr>
                        <td colSpan={accentSens ? 3 : 2} style={{ padding: '5px 6px', fontWeight: 600, fontSize: '11px', color: 'var(--text2)' }}>Total</td>
                        <td style={{ padding: '5px 6px', textAlign: 'right', fontWeight: 700, whiteSpace: 'nowrap' }}>{totalM2.toLocaleString('fr-FR')} m²</td>
                        <td />
                      </tr>
                    </tfoot>
                  )}
                </table>
                {parkRows.length > 0 && (
                  <div style={{ marginTop: '6px', display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
                    <span style={{ fontSize: '10px', fontWeight: 600, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.05em' }}>Parking</span>
                    {parkRows.map((r, i) => (
                      <span key={i} style={{ fontSize: '11px', background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: '4px', padding: '2px 7px', color: 'var(--text2)' }}>
                        {r.niveau || r.localisation || `Lot ${i+1}`}{r.surface_m2 ? ` · ${r.surface_m2} m²` : ''}
                      </span>
                    ))}
                    {parkCount > 0 && <span style={{ fontSize: '11px', fontWeight: 600 }}>{parkCount.toLocaleString('fr-FR')} m²</span>}
                  </div>
                )}
                </div>
              )
            }

            return (
              <div style={{ marginTop: '12px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '10px' }}>
                  <div className="field-lbl" style={{ margin: 0 }}>Modification des surfaces</div>
                  <span className={`pill ${lbl.cls}`} style={{ fontSize: '11px' }}>{lbl.txt}</span>
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '10px', alignItems: 'start' }}>
                  <div style={{ background: 'var(--surface)', border: '1px solid var(--border2)', borderRadius: 'var(--r)', padding: '10px 12px' }}>
                    <div style={{ fontSize: '11px', fontWeight: 700, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.06em', marginBottom: '8px' }}>Assiette initiale</div>
                    <SurfMiniTable rows={avant} accentSens={false} />
                  </div>
                  <div style={{ background: 'var(--surface)', border: '1px solid var(--border2)', borderRadius: 'var(--r)', padding: '10px 12px' }}>
                    <div style={{ fontSize: '11px', fontWeight: 700, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.06em', marginBottom: '8px' }}>Locaux complémentaires</div>
                    <SurfMiniTable rows={delta} accentSens={true} />
                  </div>
                  <div style={{ background: 'var(--accent-bg)', border: '1px solid rgba(26,95,168,.15)', borderRadius: 'var(--r)', padding: '10px 12px' }}>
                    <div style={{ fontSize: '11px', fontWeight: 700, color: 'var(--accent)', textTransform: 'uppercase', letterSpacing: '.06em', marginBottom: '8px' }}>Assiette post-avenant</div>
                    <SurfMiniTable rows={apres} accentSens={false} />
                  </div>
                </div>
              </div>
            )
          })()}
        </div>
      )}
      {/* Loyer */}
      {(show('loyer_signature_montant') || (!isAv && d.surfaces_detail?.length > 0)) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Loyer</div></div>
          {d.loyer_signature_montant && (
            <div className="loyer-hero">
              <div>
                <div className="loyer-lbl">Loyer HT/HC annuel à la signature</div>
                <div style={{ display: 'flex', alignItems: 'baseline', gap: '12px', flexWrap: 'wrap' }}>
                  <div className="loyer-amount">{fmtEur(d.loyer_signature_montant) || d.loyer_signature_montant}</div>
                  {d.indexation_indice && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <span className="pill pill-blue" style={{ fontSize: '11px' }}>{d.indexation_indice}</span>
                      {inseeLoading && <span style={{ fontSize: '11px', color: 'var(--text3)', fontStyle: 'italic' }}>chargement INSEE…</span>}
                      {inseeIndex && !inseeLoading && (
                        <span style={{ fontSize: '12px', color: 'var(--text2)', background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: '6px', padding: '2px 8px' }}>
                          {inseeIndex.label && <span style={{ fontWeight: 600, marginRight: '4px' }}>{inseeIndex.label}</span>}
                          <span>{inseeIndex.value?.toLocaleString('fr-FR', { minimumFractionDigits: 1, maximumFractionDigits: 2 })}</span>
                          <span style={{ marginLeft: '4px', fontSize: '10px', color: inseeIndex.source === 'bail' ? 'var(--success)' : 'var(--text3)', fontStyle: 'italic' }}>
                            ({inseeIndex.source === 'bail' ? 'bail' : 'INSEE'})
                          </span>
                        </span>
                      )}
                    </div>
                  )}
                </div>
              </div>
              {pills.length > 0 && <div className="pills">{pills.map((p, i) => <span key={i} className={`pill ${p.cls}`}>{p.label}</span>)}</div>}
            </div>
          )}
          {!isAv && d.surfaces_detail?.length > 0 && (() => {
            const enriched = computeUnitPrices(d.surfaces_detail, d.parking_nb_places, null)
            return (
              <div style={{ marginBottom: '16px' }}>
                <div className="field-lbl" style={{ marginBottom: '6px' }}>Ventilation du loyer par composante</div>
                <SurfaceTable surfaces={enriched} />
              </div>
            )
          })()}
          <div className="gx" style={{ marginBottom: '8px' }}>
            {show('loyer_cours') && d.loyer_cours && (() => {
              const amtCours = parseAmount(d.loyer_cours)
              const amtSig   = parseAmount(d.loyer_signature_montant)
              // Ne montrer que si différent du loyer à la signature
              if (amtCours !== null && amtSig !== null && Math.abs(amtCours - amtSig) < 1) return null
              const suspicious = amtCours !== null && amtCours < 5000
              return (
                <div className="field">
                  <div className="field-lbl">Loyer de base (annuel)</div>
                  <div className="field-val" style={{ color: suspicious ? 'var(--danger)' : undefined }}>
                    {fmtEur(d.loyer_cours) || d.loyer_cours}
                    {suspicious && <span style={{ fontSize: '11px', marginLeft: '6px', fontWeight: 400 }}>⚠ Valeur anormalement basse — vérifier (prix/m² au lieu du total ?)</span>}
                  </div>
                </div>
              )
            })()}
          </div>
          {show('indexation') && <Field label="Indexation / indice" value={d.indexation} verbose full />}
          {show('loyer_signature') && d.loyer_signature && (
            <div style={{ marginBottom: '8px' }}>
              <Field label="Loyer à la signature — détail complet" value={d.loyer_signature} verbose />
            </div>
          )}
          {(d.franchise_periodes?.length > 0 || d.franchise) && (
            <div style={{ marginTop: '8px' }}>
              {d.franchise_periodes?.length > 0 && (
                <div style={{ marginBottom: '4px' }}>
                  <div className="field-lbl" style={{ marginBottom: '6px', marginTop: '24px' }}>Franchise — périodes</div>
                  <FranchiseTable periodes={d.franchise_periodes} />
                </div>
              )}
              {d.franchise && <Field label="Franchise — modalités complètes" value={d.franchise} verbose />}
            </div>
          )}
          <div className="gx" style={{ marginTop: '8px' }}>
          </div>
        </div>
      )}

      {/* Charges / TEOM */}
      {show('charges') && d.charges && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Charges / TEOM</div></div>
          <BulletField label="" value={d.charges} full />
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
      {(show('travaux_montant') || (d.participations_travaux?.length > 0)) && (d.travaux_montant || d.travaux_modalites || d.participations_travaux?.length > 0) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Participation travaux bailleur</div></div>
          {d.participations_travaux?.length > 0 ? (
            <div className="table-wrap" style={{ marginBottom: '8px' }}>
              <table className="indemnites-table">
                <thead>
                  <tr>
                    <th>Locaux / Lot</th>
                    <th style={{ textAlign: 'right' }}>Montant max. HT</th>
                    <th>Date limite factures</th>
                    <th>Remarque</th>
                  </tr>
                </thead>
                <tbody>
                  {d.participations_travaux.map((row, i) => (
                    <tr key={i}>
                      <td style={{ fontWeight: 500, minWidth: '200px' }}>{safeStr(row.libelle) || '—'}</td>
                      <td style={{ textAlign: 'right', fontWeight: 600 }}>{row.montant ? fmtEur(row.montant) : '—'}</td>
                      <td style={{ color: 'var(--text2)' }}>{safeStr(row.date_limite) || '—'}</td>
                      <td style={{ color: 'var(--text2)', fontStyle: row.remarque ? 'normal' : 'italic' }}>{safeStr(row.remarque) || '—'}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr style={{ borderTop: '1px solid var(--border2)' }}>
                    <td style={{ fontWeight: 700, padding: '6px 10px' }}>Total</td>
                    <td style={{ textAlign: 'right', fontWeight: 700, padding: '6px 10px' }}>
                      {fmtEur(d.participations_travaux.reduce((acc, r) => acc + (parseAmount(r.montant) || 0), 0))}
                    </td>
                    <td colSpan={2} />
                  </tr>
                </tfoot>
              </table>
            </div>
          ) : (
            <div className="gx" style={{ marginBottom: '8px' }}>
              <div className="field">
                <div className="field-lbl">Montant</div>
                <div className={`field-val${!d.travaux_montant ? ' empty' : ''}`} style={d.travaux_montant ? { fontWeight: 600 } : {}}>
                  {d.travaux_montant ? (fmtEur(d.travaux_montant) || d.travaux_montant) : 'Non renseigné'}
                </div>
              </div>
              {show('travaux_date_factures') && <Field label="Date limite réception factures" value={d.travaux_date_factures} />}
            </div>
          )}
          {show('travaux_modalites') && d.travaux_modalites && (
            <Field label="Modalités complètes" value={d.travaux_modalites} verbose />
          )}
        </div>
      )}

      {/* Paliers de loyer */}
      {d.paliers_loyer?.length > 0 && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Paliers de loyer</div></div>
          <div className="table-wrap">
            <table className="indemnites-table">
              <thead><tr>
                <th>Date début</th><th>Date fin</th>
                <th style={{ textAlign: 'right' }}>Montant annuel HT/HC</th>
                <th>Description</th>
              </tr></thead>
              <tbody>
                {d.paliers_loyer.map((row, i) => (
                  <tr key={i}>
                    <td>{safeStr(row.date_debut) || '—'}</td>
                    <td>{safeStr(row.date_fin) || '—'}</td>
                    <td style={{ textAlign: 'right', fontWeight: 600 }}>{row.montant ? fmtEur(row.montant) : '—'}</td>
                    <td style={{ color: 'var(--text2)' }}>{safeStr(row.description) || '—'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Abattements temporaires */}
      {d.abattements?.length > 0 && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Abattements et réductions temporaires de loyer</div></div>
          <div className="table-wrap">
            <table className="indemnites-table">
              <thead><tr>
                <th>Date début</th><th>Date fin</th>
                <th style={{ textAlign: 'right' }}>Montant annuel HT/HC</th>
                <th>Description</th>
              </tr></thead>
              <tbody>
                {d.abattements.map((row, i) => (
                  <tr key={i}>
                    <td>{safeStr(row.date_debut) || '—'}</td>
                    <td>{safeStr(row.date_fin) || '—'}</td>
                    <td style={{ textAlign: 'right', fontWeight: 600 }}>{row.montant_annuel ? fmtEur(row.montant_annuel) : '—'}</td>
                    <td style={{ color: 'var(--text2)' }}>{safeStr(row.description) || '—'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Loyer variable */}
      {d.loyer_variable && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Loyer variable</div></div>
          <div className="gx" style={{ marginBottom: '8px' }}>
            {d.loyer_variable.type && <Field label="Type" value={safeStr(d.loyer_variable.type)} />}
            {d.loyer_variable.taux && <Field label="Taux" value={safeStr(d.loyer_variable.taux)} />}
            {d.loyer_variable.assiette && <Field label="Assiette de calcul" value={safeStr(d.loyer_variable.assiette)} />}
            {d.loyer_variable.plancher && <Field label="Plancher" value={fmtEur(d.loyer_variable.plancher) || safeStr(d.loyer_variable.plancher)} />}
            {d.loyer_variable.plafond && <Field label="Plafond" value={fmtEur(d.loyer_variable.plafond) || safeStr(d.loyer_variable.plafond)} />}
          </div>
          {d.loyer_variable.regles && (
            <Field label="Formule et règles de déclenchement" value={safeStr(d.loyer_variable.regles)} verbose />
          )}
        </div>
      )}

      {/* Indemnités par terme de bail */}
      {d.indemnites_restitution?.length > 0 && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Indemnités par terme (breaks & fin de bail)</div></div>
          <div className="table-wrap">
            <table className="indemnites-table">
              <thead><tr>
                <th>Terme / Échéance</th>
                <th>Due par</th>
                <th>Motif</th>
                <th style={{ textAlign: 'right' }}>Montant</th>
                <th>Base de calcul</th>
              </tr></thead>
              <tbody>
                {d.indemnites_restitution.map((row, i) => (
                  <tr key={i}>
                    <td style={{ fontWeight: 500, whiteSpace: 'nowrap' }}>{safeStr(row.terme) || '—'}</td>
                    <td>
                      {row.due_par && (
                        <span className={`pill ${row.due_par === 'Preneur' ? 'pill-danger' : 'pill-blue'}`} style={{ fontSize: '11px' }}>
                          {row.due_par}
                        </span>
                      )}
                    </td>
                    <td style={{ fontWeight: 500 }}>{safeStr(row.motif) || '—'}</td>
                    <td style={{ textAlign: 'right' }}>{row.montant ? fmtEur(row.montant) : '—'}</td>
                    <td style={{ color: 'var(--text2)' }}>{safeStr(row.calcul) || '—'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}


      {d.indemnites_break?.length > 0 && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Indemnités dues par le preneur en cas d'exercice d'une option de break</div></div>
          <div className="table-wrap">
            <table className="indemnites-table">
              <thead><tr>
                <th>Date de break</th>
                <th>Motif</th>
                <th style={{ textAlign: 'right' }}>Montant</th>
                <th>Base de calcul / Formule</th>
              </tr></thead>
              <tbody>
                {d.indemnites_break.map((row, i) => (
                  <tr key={i}>
                    <td style={{ whiteSpace: 'nowrap' }}>{safeStr(row.break_date) || '—'}</td>
                    <td>{safeStr(row.motif) || '—'}</td>
                    <td style={{ textAlign: 'right' }}>{row.montant ? fmtEur(row.montant) : '—'}</td>
                    <td style={{ color: 'var(--text2)' }}>{safeStr(row.calcul) || '—'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Jouissance */}
      {(show('destination') || show('article_606')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Refacturation et jouissance</div></div>
          <div className="gx">
            {show('destination') && <Field label="Destination" value={d.destination} verbose />}
            {show('article_606') && <Field label="Article 606" value={d.article_606} verbose />}
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

function ConfirmModal({ title, message, confirmLabel, onConfirm, onCancel, danger }) {
  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal" style={{ maxWidth: '400px' }} onClick={e => e.stopPropagation()}>
        <div style={{ display: 'flex', alignItems: 'flex-start', gap: '14px' }}>
          <div style={{ width: '36px', height: '36px', borderRadius: '50%', background: danger ? 'var(--danger-bg)' : 'var(--accent-bg)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={danger ? 'var(--danger)' : 'var(--accent)'} strokeWidth="2">
              <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
              <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
            </svg>
          </div>
          <div>
            <div style={{ fontWeight: 700, fontSize: '15px', marginBottom: '6px' }}>{title}</div>
            <div style={{ fontSize: '13px', color: 'var(--text2)', lineHeight: '1.5' }}>{message}</div>
          </div>
        </div>
        <div style={{ display: 'flex', gap: '8px', justifyContent: 'flex-end', marginTop: '4px' }}>
          <button className="btn" onClick={onCancel}>Annuler</button>
          <button
            className="btn"
            style={{ background: danger ? 'var(--danger)' : 'var(--accent)', color: '#fff', border: 'none' }}
            onClick={onConfirm}
          >{confirmLabel || 'Confirmer'}</button>
        </div>
      </div>
    </div>
  )
}

function Dashboard({ tree, onSelect, onDelete, onClear, onExportAll, newIds, onRefresh, onUpdateActif }) {
  const [filter, setFilter] = useState('all')
  const [confirmClear, setConfirmClear] = useState(false)
  const [confirmDelete, setConfirmDelete] = useState(null) // item to delete

  // Flatten all items for table
  const [expanded, setExpanded] = useState({})
  const [search, setSearch] = useState('')
  const [sortDir, setSortDir] = useState('asc')
  const [editingActif, setEditingActif] = useState(null) // { id, value }

  const savingRef = useRef(false)

  async function saveActifGroup(id, value) {
    if (savingRef.current) return
    savingRef.current = true
    const v = (value || '').trim()
    setEditingActif(null)
    // Update locally immediately
    onUpdateActif?.(id, v)
    // Persist to Supabase
    const { error } = await supabase
      .from('extractions')
      .update({ actif_group: v || null })
      .eq('id', id)
    if (!error) {
      await supabase.from('extractions').update({ actif_group: v || null }).eq('parent_id', id)
    }
    savingRef.current = false
  }

  function toggleExpand(id) {
    setExpanded(prev => ({ ...prev, [id]: !prev[id] }))
  }

  // Build display rows based on filter and expanded state
  const displayRows = []
  tree.forEach(bail => {
    const bailRow = { ...bail, _level: 0, _parentName: null, _bailData: bail.data }
    // Trier avenants par date de signature puis numéroter
    function toSortable(dateStr) {
      if (!dateStr) return ''
      const s = String(dateStr)
      // dd/mm/yyyy → yyyy-mm-dd
      const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/)
      if (m) return `${m[3]}-${m[2]}-${m[1]}`
      return s
    }
    const sortedAv = [...(bail.avenants || [])].sort((a, b) => {
      const da = toSortable(a.data?.date_effet_avenant || a.data?.date_signature_avenant || a.created_at || '')
      const db = toSortable(b.data?.date_effet_avenant || b.data?.date_signature_avenant || b.created_at || '')
      return da.localeCompare(db)
    })
    const avRows = sortedAv.map((av, idx) => ({
      ...av, _level: 1,
      _parentName: bail.data?.immeuble || bail.data?.adresse || bail.file_name,
      _bailData: bail.data,
      _avNum: idx + 1
    }))
    if (filter === 'avenant') {
      avRows.forEach(r => displayRows.push(r))
      return
    }
    displayRows.push({ ...bailRow, _avCount: avRows.length })
    if (filter !== 'bail' && expanded[bail.id] && avRows.length > 0) {
      avRows.forEach(r => displayRows.push(r))
    }
  })
  if (filter === 'avenant') {
    // orphan avenants
    tree.filter(r => r.document_type === 'avenant').forEach(av => {
      displayRows.push({ ...av, _level: 0, _parentName: null, _bailData: null })
    })
  }

  // Apply search filter
  const q = search.trim().toLowerCase()
  const filtered = q ? displayRows.filter(row => {
    const d = row.data || {}
    const searchIn = [
      d.immeuble, d.adresse, d.ville, d.preneur, d.bailleur,
      d.objet_avenant, row.file_name, row._parentName,
    ].filter(Boolean).join(' ').toLowerCase()
    return searchIn.includes(q)
  }) : displayRows

  // Sort top-level bails by actif name, avenants follow their bail
  const getActifName = row => (row.data?.immeuble || row.data?.adresse || row.file_name || '').toLowerCase()
  const sortedFiltered = (() => {
    const bails = filtered.filter(r => r._level === 0)
    const avMap = {}
    filtered.filter(r => r._level > 0).forEach(r => {
      const pid = r.parent_id || r._parentId
      if (!avMap[pid]) avMap[pid] = []
      avMap[pid].push(r)
    })
    bails.sort((a, b) => {
      const cmp = getActifName(a).localeCompare(getActifName(b), 'fr')
      return sortDir === 'asc' ? cmp : -cmp
    })
    // Group by actif_group
    const groups = {} // actif_group -> [bail rows + avenants]
    const noGroup = [] // bails without actif_group
    bails.forEach(bail => {
      const grp = bail.actif_group
      const rows = [bail, ...(avMap[bail.id] || [])]
      if (grp) {
        if (!groups[grp]) groups[grp] = []
        groups[grp].push(...rows)
      } else {
        noGroup.push(...rows)
      }
    })
    // Build final list: group headers + rows
    const result = []
    Object.entries(groups).sort(([a], [b]) => {
      const cmp = a.localeCompare(b, 'fr')
      return sortDir === 'asc' ? cmp : -cmp
    }).forEach(([grp, rows]) => {
      result.push({ _isGroupHeader: true, _groupName: grp, _groupCount: rows.filter(r => r._level === 0).length })
      rows.forEach(r => result.push(r))
    })
    noGroup.forEach(r => result.push(r))
    return result
  })()

  return (
    <div className="dashboard">
      {confirmClear && (
        <ConfirmModal
          title="Vider le dashboard ?"
          message="Toutes les extractions seront supprimées définitivement. Cette action est irréversible."
          confirmLabel="Vider"
          danger
          onConfirm={() => { setConfirmClear(false); onClear() }}
          onCancel={() => setConfirmClear(false)}
        />
      )}
      {confirmDelete && (
        <ConfirmModal
          title="Supprimer cette extraction ?"
          message={`"${confirmDelete.data?.immeuble || confirmDelete.data?.adresse || confirmDelete.file_name}" sera supprimé définitivement.`}
          confirmLabel="Supprimer"
          danger
          onConfirm={e => { onDelete(confirmDelete, { stopPropagation: () => {} }); setConfirmDelete(null) }}
          onCancel={() => setConfirmDelete(null)}
        />
      )}
      {/* Toolbar */}
      <div className="dash-toolbar">
        <div className="dash-stats">
          <span className="dash-stat">{tree.length} {tree.length !== 1 ? 'baux' : 'bail'}</span>
          <span className="dash-stat">{tree.reduce((a,b) => a + (b.avenants?.length||0), 0)} avenant{tree.reduce((a,b) => a + (b.avenants?.length||0), 0) !== 1 ? 's' : ''}</span>
        </div>
        <div style={{ flex: 1, maxWidth: '320px', position: 'relative' }}>
          <input
            type="text"
            value={search}
            onChange={e => setSearch(e.target.value)}
            placeholder="Rechercher actif, preneur, ville…"
            style={{
              width: '100%', boxSizing: 'border-box',
              padding: '6px 30px 6px 10px',
              fontSize: '13px', border: '1px solid var(--border2)',
              borderRadius: '6px', background: 'var(--surface)',
              color: 'var(--text)', outline: 'none',
            }}
          />
          {search && (
            <button onClick={() => setSearch('')} style={{
              position: 'absolute', right: '8px', top: '50%', transform: 'translateY(-50%)',
              background: 'none', border: 'none', cursor: 'pointer',
              color: 'var(--text3)', fontSize: '14px', padding: '0', lineHeight: 1,
            }}>✕</button>
          )}
        </div>
        <div className="dash-filters">
          {['all','bail','avenant'].map(f => (
            <button key={f} className={`dash-filter${filter === f ? ' active' : ''}`} onClick={() => setFilter(f)}>
              {f === 'all' ? 'Tous' : f === 'bail' ? 'Baux' : 'Avenants'}
            </button>
          ))}
        </div>
        {tree.length > 0 && (
          <div style={{ display: 'flex', gap: '6px' }}>
            <button className="btn" style={{ width: 'auto', padding: '5px 12px', display: 'flex', alignItems: 'center', gap: '5px' }} onClick={onExportAll}>
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
              </svg>
              Exporter tout
            </button>
            <button className="btn-clear" style={{ width: 'auto', padding: '5px 12px' }} onClick={() => setConfirmClear(true)}>Vider</button>
          </div>
        )}
      </div>

      {/* Table */}
      {!filtered.length ? (
        <div className="dash-empty">
          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.2" style={{ color: 'rgba(255,255,255,0.2)', marginBottom: '10px' }}>
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
            <polyline points="14 2 14 8 20 8"/>
          </svg>
          <div>Aucune extraction</div>
        </div>
      ) : (
        <div className="dash-table">
          <div className="dash-thead">
            <div className="dash-th" style={{ gridColumn: '1', cursor: 'pointer', userSelect: 'none', display: 'flex', alignItems: 'center', gap: '5px' }}
              onClick={() => setSortDir(d => d === 'asc' ? 'desc' : 'asc')}>
              Actif / Document
              <span style={{ fontSize: '10px', color: 'var(--text3)' }}>{sortDir === 'asc' ? '↑' : '↓'}</span>
            </div>
            <div className="dash-th" style={{ gridColumn: '2' }}>Type</div>
            <div className="dash-th" style={{ gridColumn: '3' }}>Preneur</div>
            <div className="dash-th" style={{ gridColumn: '4' }}>Date effet</div>
            <div className="dash-th" style={{ gridColumn: '5' }}>Date fin</div>
            <div className="dash-th" style={{ gridColumn: '6' }}>Break</div>
            <div className="dash-th dash-th-right" style={{ gridColumn: '7' }}>Loyer HT/HC</div>
            <div style={{ gridColumn: '8' }}/>
          </div>
          {sortedFiltered.map((row, rowIdx) => {
            // Group header
            if (row._isGroupHeader) return (
              <div key={`grp-${row._groupName}`} style={{
                display: 'grid', gridTemplateColumns: '1fr auto',
                alignItems: 'center', padding: '10px 16px 6px',
                background: 'var(--surface2)', borderBottom: '1px solid var(--border2)',
                marginTop: rowIdx > 0 ? '4px' : 0,
              }}>
                <span style={{ fontSize: '11px', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '.08em', color: 'var(--accent)' }}>
                  📁 {row._groupName}
                </span>
                <span style={{ fontSize: '11px', color: 'var(--text3)' }}>{row._groupCount} bail{row._groupCount > 1 ? 's' : ''}</span>
              </div>
            )

            // Données fusionnées : bail de base + modifications de l'avenant
            const bailBase = row._bailData || {}
            const mods = row.data?.champs_modifies || {}
            const d = row.document_type === 'avenant'
              ? {
                  // Hériter les valeurs du bail parent
                  immeuble: bailBase.immeuble,
                  adresse: bailBase.adresse,
                  ville: bailBase.ville,
                  preneur: bailBase.preneur,
                  loyer_signature_montant: mods.loyer_signature_montant ?? null,
                  date_effet: row.data?.date_effet_avenant || mods.date_effet || null,
                  date_fin: mods.date_fin || bailBase.date_fin,
                  break_options: mods.break_options || bailBase.break_options,
                  objet_avenant: row.data?.objet_avenant,
                }
              : (row.data || {})
            const isNew = newIds?.includes(row.id)
            const isAv = row.document_type === 'avenant'
            const breaks = Array.isArray(d.break_options) ? d.break_options : []
            return (
              <div
                key={row.id}
                className={`dash-row${isNew ? ' dash-row-new' : ''}${row._level ? ' dash-row-av' : ''}`}
                onClick={() => row._level === 0 && row._avCount > 0 ? toggleExpand(row.id) : onSelect(row)}
              >
                {/* Actif / Document */}
                <div className="dash-td" style={{ paddingLeft: row._level ? '32px' : '16px', flexDirection: 'column', alignItems: 'flex-start', gap: '2px' }}>

                  <div style={{ fontWeight: 700, fontSize: '13px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', width: '100%', color: 'var(--text)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                    {!isAv && row._avCount > 0 && (
                      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"
                        style={{ flexShrink: 0, color: 'var(--text3)', transform: expanded[row.id] ? 'rotate(90deg)' : 'rotate(0deg)', transition: 'transform .2s' }}>
                        <polyline points="9 18 15 12 9 6"/>
                      </svg>
                    )}
                    <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {isAv
                        ? `Avenant ${row._avNum || ''}`
                        : (d.immeuble || d.adresse || row.file_name.replace(/\.[^.]+$/, '')).toUpperCase()
                      }
                    </span>
                    {!isAv && row._avCount > 0 && (
                      <span style={{ fontSize: '10px', fontWeight: 600, padding: '1px 6px', borderRadius: '99px', background: 'var(--surface2)', color: 'var(--text3)', flexShrink: 0 }}>
                        {row._avCount} av.
                      </span>
                    )}
                  </div>
                  <div style={{ fontSize: '11px', color: 'var(--text3)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', width: '100%' }}>
                    {isAv ? (row._parentName || '') : (() => {
                      const v = d.ville || ''
                      // Extract postal code if present: "Aix-en-Provence (13290)" or "13290 Aix"
                      const cpMatch = v.match(/(\d{5})/)
                      const cp = cpMatch ? cpMatch[1] : (d.adresse?.match(/(\d{5})/)?.[1] || '')
                      const cityOnly = v.replace(/\d{5}\s*/g, '').replace(/[()]/g, '').trim()
                      const city = cityOnly.toUpperCase()
                      return cp ? `${city} (${cp})` : city
                    })()}
                  </div>
                  {!isAv && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '4px', marginTop: '2px' }}>
                      {editingActif?.id === row.id ? (
                        <input
                          autoFocus
                          defaultValue={editingActif.value}
                          placeholder="Nom de l'actif groupant…"
                          onClick={e => e.stopPropagation()}
                          onKeyDown={e => {
                            e.stopPropagation()
                            if (e.key === 'Enter') { e.preventDefault(); saveActifGroup(row.id, e.target.value) }
                            if (e.key === 'Escape') { e.preventDefault(); setEditingActif(null) }
                          }}
                          onBlur={e => saveActifGroup(row.id, e.target.value)}
                          style={{ fontSize: '10px', padding: '1px 5px', border: '1px solid var(--accent)', borderRadius: '4px', outline: 'none', width: '140px', background: 'var(--surface)', color: 'var(--text)' }}
                        />
                      ) : (
                        <span
                          onClick={e => { e.stopPropagation(); setEditingActif({ id: row.id, value: row.actif_group || '' }) }}
                          title="Cliquer pour définir l'actif groupant"
                          style={{ fontSize: '10px', padding: '1px 6px', borderRadius: '4px', cursor: 'pointer',
                            background: row.actif_group ? 'var(--accent-bg)' : 'var(--surface2)',
                            color: row.actif_group ? 'var(--accent)' : 'var(--text3)',
                            border: `1px solid ${row.actif_group ? 'rgba(26,95,168,.2)' : 'var(--border)'}`,
                            fontWeight: row.actif_group ? 600 : 400,
                          }}>
                          {row.actif_group || '+ Actif'}
                        </span>
                      )}
                    </div>
                  )}
                </div>

                {/* Type */}
                <div className="dash-td">
                  <span className={`dash-tag ${isAv ? 'dash-tag-av' : 'dash-tag-bail'}`}>
                    {isAv ? 'Avenant' : 'Bail'}
                  </span>
                </div>

                {/* Preneur */}
                <div className="dash-td">
                  <span style={{ fontSize: '12px', color: 'var(--text)', fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', display: 'block' }}>
                    {shortPartyName(d.preneur)?.toUpperCase() || '—'}
                  </span>
                </div>

                {/* Date effet */}
                <div className="dash-td">
                  <span style={{ fontSize: '12px', color: 'var(--text2)' }}>{normalizeDate(d.date_effet) || '—'}</span>
                </div>

                {/* Date fin */}
                <div className="dash-td">
                  <span style={{ fontSize: '12px', color: 'var(--text2)' }}>{normalizeDate(d.date_fin) || '—'}</span>
                </div>

                {/* Break */}
                <div className="dash-td">
                  {breaks.length > 0 ? (
                    <div style={{ display: 'flex', gap: '3px', flexWrap: 'wrap' }}>
                      {breaks
                        .filter(b => typeof b === 'string' && b.length < 30) // exclure texte verbeux
                        .slice(0, 2).map((b, i) => (
                          <span key={i} style={{ fontSize: '10px', fontWeight: 600, padding: '1px 6px', borderRadius: '3px', background: 'var(--accent-bg)', color: 'var(--accent)', border: '1px solid rgba(26,95,168,0.2)', whiteSpace: 'nowrap' }}>{normalizeDate(b) || b}</span>
                        ))}
                      {breaks.filter(b => typeof b === 'string' && b.length < 30).length > 2 && <span style={{ fontSize: '10px', color: 'var(--text3)' }}>+{breaks.length-2}</span>}
                    </div>
                  ) : <span style={{ fontSize: '12px', color: 'var(--text3)' }}>—</span>}
                </div>

                {/* Loyer */}
                <div className="dash-td dash-td-right">
                  {d.loyer_signature_montant ? (
                    <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text)', letterSpacing: '-0.01em' }}>
                      {fmtEur(d.loyer_signature_montant)}
                    </span>
                  ) : <span style={{ fontSize: '12px', color: 'var(--text3)' }}>—</span>}
                </div>

                {/* Actions */}
                <div className="dash-td dash-td-actions" onClick={e => e.stopPropagation()}>
                  <button className="dash-action-btn" style={{ opacity: 1 }} onClick={e => { e.stopPropagation(); onSelect(row) }} title="Voir le détail">
                    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
                  </button>
                  <button className="dash-action-btn dash-action-del" onClick={e => { e.stopPropagation(); setConfirmDelete(row) }} title="Supprimer">
                    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/></svg>
                  </button>
                </div>
              </div>
            )
          })}
        </div>
      )}
    </div>
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
  const [lastError,    setLastError]    = useState('')
  const [newIds,       setNewIds]       = useState([])   // ids extraits dans le batch courant

  function buildTree(rows) {
    const bails    = rows.filter(r => r.document_type === 'bail')
    const avenants = rows.filter(r => r.document_type === 'avenant' && r.parent_id)
    const orphans  = rows.filter(r => r.document_type === 'avenant' && !r.parent_id)
    return [...bails.map(b => ({ ...b, avenants: avenants.filter(a => a.parent_id === b.id) })), ...orphans]
  }

  async function loadHistory() {
    if (histLoaded) return
    const { data: rows } = await supabase.from('extractions')
      .select('id, file_name, created_at, data, document_type, parent_id, actif_group')
      .order('created_at', { ascending: false }).limit(100)
    setHistory(rows ? buildTree(rows) : [])
    setHistLoaded(true)
  }

  async function switchTab(t) {
    setTab(t)
    if (t === 'history') {
      // Forcer rechargement depuis Supabase directement
      const { data: rows } = await supabase.from('extractions')
        .select('id, file_name, created_at, data, document_type, parent_id, actif_group')
        .order('created_at', { ascending: false }).limit(100)
      setHistory(rows ? buildTree(rows) : [])
      setHistLoaded(true)
    }
  }
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
          const promptWithName = DETECT_PROMPT + `\n\nNom du fichier: "${newFiles[i].name}"`
          const data = await callClaude(base64, mediaType, promptWithName)
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
    // Pré-remplir bail lié : baux en base + baux du batch (id virtuel batch-N)
    const existingBails = history.filter(h => h.document_type === 'bail')
    const batchBails = bailIdx
      .filter(i => pertinents[i] !== false) // exclure non pertinents
      .map(i => ({
        id: `batch-${i}`,
        file_name: newFiles[i].name,
        data: { preneur: detectData[i]?.preneur, bailleur: detectData[i]?.bailleur, adresse: detectData[i]?.adresse, immeuble: detectData[i]?.immeuble }
      }))
    const allBailsForMatch = [...existingBails, ...batchBails]
    const autoLinks = {}
    avenantIdx
      .filter(i => pertinents[i] !== false) // exclure non pertinents
      .forEach(i => {
        if (allBailsForMatch.length === 1) {
          autoLinks[i] = allBailsForMatch[0].id
        } else if (allBailsForMatch.length > 1 && detectData[i]) {
          const match = findBestMatch(detectData[i], allBailsForMatch)
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
    // Exclure les fichiers non pertinents
    const pertinent = (i) => pertinents[i] !== false
    const bailIndices    = order.filter(i => (docTypes[i] || 'bail') === 'bail' && pertinent(i))
    const avenantIndices = order.filter(i => docTypes[i] === 'avenant' && pertinent(i))
    const availableBails = [...history.filter(h => h.document_type === 'bail')]

    // 1. Extraire les baux d'abord
    for (const i of bailIndices) {
      try {
        setStatus(i, 'loading')
        const base64 = await toBase64(files[i])
        const mediaType = getMediaType(files[i])
        const extracted = await callClaude(base64, mediaType, EXTRACTION_PROMPT)
        // Appels dédiés en parallèle : breaks + financier critique
        try {
          const [breakResult, financialResult] = await Promise.all([
            callClaude(base64, mediaType, BREAK_PROMPT).catch(() => null),
            callClaude(base64, mediaType, FINANCIAL_PROMPT).catch(() => null),
          ])
          if (breakResult?.break_options?.length > 0) extracted.break_options = breakResult.break_options
          if (breakResult?.date_fin && !extracted.date_fin) extracted.date_fin = breakResult.date_fin
          if (financialResult) {
            const f = financialResult
            if (f.loyer_signature_montant) extracted.loyer_signature_montant = f.loyer_signature_montant
            if (f.loyer_signature) extracted.loyer_signature = f.loyer_signature
            if (Array.isArray(f.franchise_periodes) && f.franchise_periodes.length > 0) extracted.franchise_periodes = f.franchise_periodes
            if (Array.isArray(f.participations_travaux) && f.participations_travaux.length > 0) extracted.participations_travaux = f.participations_travaux
            if (Array.isArray(f.paliers_loyer) && f.paliers_loyer.length > 0) extracted.paliers_loyer = f.paliers_loyer
            if (Array.isArray(f.abattements) && f.abattements.length > 0) extracted.abattements = f.abattements
            if (f.loyer_variable) extracted.loyer_variable = f.loyer_variable
            if (Array.isArray(f.indemnites_break) && f.indemnites_break.length > 0) extracted.indemnites_break = f.indemnites_break
          }
        } catch (_) { /* non bloquant */ }
        const saved = await saveExtraction(files[i], extracted, 'bail', null)
        if (saved) {
          const bwa = { ...saved, avenants: [] }
          availableBails.push(bwa)
          setNewIds(prev => [...prev, saved.id])
          // Ne pas setHistory ici — on recharge tout à la fin
        }
        setStatus(i, 'done')
      } catch (e) { setStatus(i, 'error', e.message); setLastError(e.message) }
    }

    // 2. Extraire les avenants et sauvegarder directement avec le bail lié choisi
    let lastSaved = null
    for (const i of avenantIndices) {
      try {
        setStatus(i, 'loading')
        const base64 = await toBase64(files[i])
        const mediaType = getMediaType(files[i])
        const extracted = await callClaude(base64, mediaType, AVENANT_PROMPT)
        // Appel dédié financier pour les avenants
        try {
          const financialResult = await callClaude(base64, mediaType, FINANCIAL_PROMPT).catch(() => null)
          if (financialResult) {
            const f = financialResult
            const mods = extracted.champs_modifies || {}
            if (f.loyer_signature_montant) mods.loyer_signature_montant = f.loyer_signature_montant
            if (f.loyer_signature) mods.loyer_signature = f.loyer_signature
            if (Array.isArray(f.franchise_periodes) && f.franchise_periodes.length > 0) mods.franchise_periodes = f.franchise_periodes
            if (Array.isArray(f.participations_travaux) && f.participations_travaux.length > 0) mods.participations_travaux = f.participations_travaux
            if (Array.isArray(f.paliers_loyer) && f.paliers_loyer.length > 0) mods.paliers_loyer = f.paliers_loyer
            if (Array.isArray(f.abattements) && f.abattements.length > 0) mods.abattements = f.abattements
            if (f.loyer_variable) mods.loyer_variable = f.loyer_variable
            if (Array.isArray(f.indemnites_break) && f.indemnites_break.length > 0) mods.indemnites_break = f.indemnites_break
            extracted.champs_modifies = mods
          }
        } catch (_) { /* non bloquant */ }
        // Résoudre batch- id en vrai id
        let parentId = avenantLinks[i] || null
        if (parentId && parentId.startsWith('batch-')) {
          const batchIdx = parseInt(parentId.replace('batch-', ''))
          const realBail = availableBails.find(b => b.file_name === files[batchIdx]?.name)
          parentId = realBail?.id || null
        }
        const saved = await saveExtraction(files[i], extracted, 'avenant', parentId)
        if (saved) {
          lastSaved = saved
          setNewIds(prev => [...prev, saved.id])
          // Ne pas setHistory ici — on recharge tout à la fin
        }
        setStatus(i, 'done')
      } catch (e) { setStatus(i, 'error', e.message); setLastError(e.message) }
    }

    setLoading(false)
    // Recharger l'historique complet depuis Supabase
    setHistLoaded(false)
    const { data: freshRows } = await supabase.from('extractions')
      .select('id, file_name, created_at, data, document_type, parent_id, actif_group')
      .order('created_at', { ascending: false }).limit(100)
    if (freshRows) setHistory(buildTree(freshRows))
    setHistLoaded(true)
    setTab('history')
  }


  async function handleDeleteItem(item, e) {
    e.stopPropagation()
    // Si bail : supprimer aussi les avenants liés en base
    if (item.document_type === 'bail') {
      await supabase.from('extractions').delete().eq('parent_id', item.id)
    }
    await supabase.from('extractions').delete().eq('id', item.id)
    if (activeItem?.id === item.id) setActiveItem(null)
    setHistory(prev => prev.filter(b => b.id !== item.id).map(b => ({ ...b, avenants: (b.avenants || []).filter(a => a.id !== item.id) })))
  }

  async function handleClearHistory() {
    await supabase.from('extractions').delete().neq('id', '00000000-0000-0000-0000-000000000000')
    setHistory([])
    setHistLoaded(true) // marquer comme chargé pour éviter un rechargement parasite
    setActiveItem(null)
    setNewIds([])
  }

  function handleClear() {
    setFiles([]); setStatuses([]); setActiveItem(null); setDocTypes([])
    setLastError(''); setFileOrder([]); setAvenantLinks({}); setPertinents([]); setRaisons([])
  }

  const d = activeItem?.data || {}
  const resultTitle = d.immeuble || d.adresse || activeItem?.file_name || ''
  const shortName = s => s?.split(',')[0]?.split('(')[0]?.split(' SAS')[0]?.split(' SA ')[0]?.trim()
  const resultSub = [shortName(d.preneur), shortName(d.bailleur), d.date_signature ? `Signé le ${d.date_signature}` : null].filter(Boolean).join(' · ')

  return (
    <>
      <div className="app">
        <aside className="sidebar">
          <div className="sidebar-logo" style={{ cursor: 'pointer' }} onClick={() => { handleClear(); setTab('extract') }}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14 2 14 8 20 8"/>
              <line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>
            </svg>
            Lease Reader
          </div>
          <nav className="sidebar-nav">
            <button className={`nav-item${tab === 'extract' ? ' active' : ''}`} onClick={() => { handleClear(); setTab('extract') }}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
              </svg>
              Extraire
            </button>
            <button className={`nav-item${tab === 'history' || activeItem ? ' active' : ''}`} onClick={() => { setActiveItem(null); switchTab('history') }}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <polyline points="12 8 12 12 14 14"/>
                <path d="M3.05 11a9 9 0 1 0 .5-4"/><polyline points="3 3 3 7 7 7"/>
              </svg>
              Dashboard
              {history.length > 0 && <span className="badge">{history.reduce((a,b) => a + 1 + (b.avenants?.length||0), 0)}</span>}
            </button>
          </nav>

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
                <button className="btn back" onClick={() => setActiveItem(null)}>← Retour au dashboard</button>
                <button className="btn primary" onClick={() => {
                  const bailParent = history.find(b => b.avenants?.some(a => a.id === activeItem.id))
                  exportToExcel(
                    [{ item: activeItem, parentName: bailParent?.data?.immeuble || '', parentData: bailParent?.data || null }],
                    activeItem.data?.immeuble || activeItem.file_name
                  )
                }}>
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
            {activeItem ? (
              <ResultsView item={activeItem} />
            ) : tab === 'history' ? (
              <Dashboard
                tree={history}
                onSelect={item => setActiveItem(item)}
                onDelete={handleDeleteItem}
                onClear={handleClearHistory}
                onExportAll={() => exportAllToExcel(history)}
                newIds={newIds}
                onRefresh={loadHistory}
                onUpdateActif={(id, value) => {
                  setHistory(prev => prev.map(b => {
                    if (b.id === id) return { ...b, actif_group: value || null }
                    return { ...b, avenants: (b.avenants || []).map(a => a.id === id ? { ...a, actif_group: value || null } : a) }
                  }))
                }}
              />
            ) : (
              <div className="extract-wrap">

                {/* ── Queue principale ── */}
                <>
                    <DropZone onFiles={handleFiles} disabled={loading || detecting} />
                    <PageLimitWarning />

                    {files.length > 0 && (
                      <div style={{ marginTop: '10px' }}>
                        {/* En-tête colonnes */}
                        <div style={{ display: 'grid', gridTemplateColumns: '20px 1fr 100px 120px 220px 32px', gap: '8px', padding: '0 4px 6px', borderBottom: '1px solid var(--border)', marginBottom: '4px' }}>
                          <div/>
                          <div style={{ fontSize: '10px', fontWeight: 700, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.06em' }}>Fichier</div>
                          <div style={{ fontSize: '10px', fontWeight: 700, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.06em' }}>Pertinent</div>
                          <div style={{ fontSize: '10px', fontWeight: 700, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.06em' }}>Type</div>
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
                            // Baux disponibles = historique + fichiers du batch avec toggle=bail ET pertinent
                            const batchBails = files
                              .map((bf, bi) => docTypes[bi] === 'bail' && bi !== fileIdx && pertinents[bi] !== false
                                ? { id: `batch-${bi}`, file_name: bf.name, data: {} } : null)
                              .filter(Boolean)
                            const allBails = [
                              ...history.filter(h => h.document_type === 'bail'),
                              ...batchBails
                            ]
                            return (
                              <div key={fileIdx} className={`queue-item ${st.state || ''}`}
                                style={{ display: 'grid', gridTemplateColumns: '20px 1fr 100px 120px 220px 32px', gap: '8px', alignItems: 'center', padding: '8px 4px', flexWrap: 'nowrap' }}>

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

                                {/* Toggle Bail/Avenant — grisé si non pertinent */}
                                <div>
                                  {analyzing ? (
                                    <span style={{ fontSize: '11px', color: 'var(--text3)', fontStyle: 'italic' }}>Analyse…</span>
                                  ) : pertinent === false ? (
                                    <span style={{ fontSize: '11px', color: 'var(--text3)' }}>—</span>
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

                                {/* Bail lié */}
                                <div>
                                  {isAvenant && pertinent !== false ? (
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
              </div>
            )}
          </div>
        </main>
      </div>
    </>
  )
}
