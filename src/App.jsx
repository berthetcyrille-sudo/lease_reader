import { useState, useCallback, useRef } from 'react'
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

const EXTRACTION_PROMPT = `Tu es un expert en baux commerciaux français. Analyse ce document et extrait précisément les données. Réponds UNIQUEMENT en JSON valide, sans markdown ni backticks.

RÈGLES IMPORTANTES :
- "break_options" : TABLEAU de strings, une date par élément. Ex: ["31 août 2027", "31 août 2030"]. null si aucune.
- "indemnites" : TABLEAU d'objets structurés. Ex: [{"motif":"Restitution franchise","due_par":"Preneur","montant":"123 405 €","date_limite":"À la date de départ"}]. null si aucune.
- Les champs "_montant" et "_duree" contiennent uniquement la valeur brute chiffrée.
- Les champs "_modalites", "_detail", "_formulation" contiennent la clause complète.

{
  "adresse": "adresse complète",
  "immeuble": "nom de l'immeuble",
  "ville": "ville et code postal",
  "type_bail": "type de bail",
  "duree_totale": "durée totale",
  "duree_ferme": "durée ferme",
  "preneur": "nom complet du preneur avec forme juridique et SIREN",
  "bailleur": "nom complet du bailleur avec forme juridique et SIREN",
  "garant": "garant ou caution si mentionné",
  "date_effet": "date d'effet",
  "date_signature": "date de signature",
  "break_options": ["date1", "date2"],
  "notice": "durée de préavis",
  "date_conge": "date limite pour donner congé",
  "date_fin": "date d'expiration",
  "date_limite_travaux": "date limite travaux preneur",
  "conditions_break": "conditions financières et formelles d'exercice du break",
  "surface_totale_m2": "surface totale en m² — valeur brute uniquement ex: 2224.98",
  "surface_bureaux": "détail surfaces bureaux par niveau/bâtiment",
  "surface_totale": "formulation complète de la surface totale",
  "parking": "nombre et description des places",
  "rie": "RIE : oui/non et modalités",
  "autres_surfaces": "autres surfaces (archives, locaux techniques...)",
  "loyer_signature_montant": "montant brut du loyer annuel HT/HC à la signature en chiffres uniquement ex: 493621.80",
  "loyer_signature": "formulation complète du loyer à la signature avec répartition si plusieurs composantes",
  "loyer_cours": "loyer actuel avec détail si modifié par avenant",
  "indexation": "clause d'indexation complète : indice, date de référence, modalités",
  "franchise_duree": "durée brute de la franchise ex: 12 mois",
  "franchise": "modalités complètes de la franchise : durée, période, conditions de restitution",
  "charges": "répartition des charges, TEOM, provisions",
  "depot_garantie_montant": "montant brut du dépôt de garantie en chiffres uniquement ex: 123405.30",
  "depot_garantie": "modalités complètes du dépôt de garantie : montant, durée, restitution",
  "travaux_montant": "montant brut de la participation bailleur aux travaux preneur en chiffres uniquement",
  "travaux_date_factures": "date limite de réception des factures pour la participation travaux",
  "travaux_modalites": "modalités complètes de la participation travaux : conditions d'appel, justificatifs, sort en cas de non-consommation",
  "indemnites": [{"motif": "motif", "due_par": "Preneur ou Bailleur", "montant": "montant", "date_limite": "date ou condition"}],
  "indemnites_detail": "description complète de toutes les indemnités contractuelles : restitution de franchise, indemnité libératoire de remise en état, autres pénalités",
  "article_606": "qui supporte l'article 606 et dans quelles conditions",
  "conformite": "obligations de conformité",
  "accession": "clause d'accession",
  "remise_en_etat": "obligations de remise en état à la sortie",
  "maintenance": "obligations de maintenance et travaux en cours de bail",
  "destination": "destination contractuelle des locaux",
  "sous_location": "conditions de sous-location",
  "cession": "conditions de cession du bail"
}

Si une information est absente, mets null. Reprends les montants, dates et formulations exactes du document.`

const AVENANT_PROMPT = `Tu es un expert en baux commerciaux français. Ce document est un AVENANT à un bail existant.

Réponds UNIQUEMENT en JSON valide, sans markdown ni backticks.

RÈGLES :
- "break_options" dans champs_modifies : TABLEAU de strings si modifié.
- "indemnites" dans champs_modifies : TABLEAU d'objets structurés si modifié.
- Les champs "_montant" et "_duree" contiennent uniquement la valeur brute chiffrée.

{
  "bail_reference": {
    "preneur": "nom du preneur",
    "bailleur": "nom du bailleur",
    "date_bail_origine": "date du bail d'origine",
    "adresse": "adresse",
    "immeuble": "nom de l'immeuble"
  },
  "date_effet_avenant": "date d'entrée en vigueur",
  "date_signature_avenant": "date de signature",
  "objet_avenant": "résumé en 1-2 phrases",
  "champs_modifies": {
    "adresse": null, "immeuble": null, "ville": null, "type_bail": null,
    "duree_totale": null, "duree_ferme": null, "preneur": null, "bailleur": null,
    "garant": null, "date_effet": null, "date_signature": null,
    "break_options": null, "notice": null, "date_conge": null, "date_fin": null,
    "date_limite_travaux": null, "conditions_break": null,
    "surface_totale_m2": null, "surface_bureaux": null, "surface_totale": null,
    "parking": null, "rie": null, "autres_surfaces": null,
    "loyer_signature_montant": null, "loyer_signature": null, "loyer_cours": null,
    "indexation": null, "franchise_duree": null, "franchise": null, "charges": null,
    "depot_garantie_montant": null, "depot_garantie": null,
    "travaux_montant": null, "travaux_date_factures": null, "travaux_modalites": null,
    "indemnites": null, "indemnites_detail": null,
    "article_606": null, "conformite": null, "accession": null,
    "remise_en_etat": null, "maintenance": null, "destination": null,
    "sous_location": null, "cession": null
  }
}

Dans "champs_modifies", ne renseigne QUE les champs modifiés par cet avenant. Laisse null les autres.`

// ─── CSS ─────────────────────────────────────────────────────────────────────

const CSS = `
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --bg: #f5f4f1; --surface: #ffffff; --surface2: #f0efe9;
  --border: rgba(0,0,0,0.08); --border2: rgba(0,0,0,0.14);
  --text: #181816; --text2: #5a5855; --text3: #9a9895;
  --accent: #185FA5; --accent-bg: #E6F1FB; --accent-dark: #0C447C;
  --danger: #A32D2D; --danger-bg: #FCEBEB;
  --success: #0F6E56; --success-bg: #E1F5EE;
  --amber-bg: #FAEEDA; --amber: #854F0B;
  --r: 10px; --rl: 14px;
  font-family: 'Inter', system-ui, -apple-system, sans-serif;
}
body { background: var(--bg); color: var(--text); font-size: 14px; line-height: 1.6; min-height: 100vh; }
button, input { font-family: inherit; cursor: pointer; }
::-webkit-scrollbar { width: 5px; } ::-webkit-scrollbar-thumb { background: var(--border2); border-radius: 3px; }
.app { display: flex; min-height: 100vh; }

.sidebar { width: 224px; flex-shrink: 0; background: var(--surface); border-right: 0.5px solid var(--border); display: flex; flex-direction: column; }
.sidebar-logo { display: flex; align-items: center; gap: 9px; padding: 18px 16px 14px; font-size: 14px; font-weight: 600; color: var(--accent); border-bottom: 0.5px solid var(--border); }
.sidebar-nav { padding: 10px 8px; display: flex; flex-direction: column; gap: 1px; }
.nav-item { display: flex; align-items: center; gap: 9px; padding: 8px 10px; border-radius: var(--r); border: none; background: transparent; color: var(--text2); font-size: 13px; font-weight: 500; text-align: left; width: 100%; transition: background .12s, color .12s; }
.nav-item:hover { background: var(--surface2); color: var(--text); }
.nav-item.active { background: var(--accent-bg); color: var(--accent); }
.nav-item .badge { margin-left: auto; background: var(--accent-bg); color: var(--accent); font-size: 11px; font-weight: 600; padding: 1px 6px; border-radius: 999px; }
.history-list { flex: 1; overflow-y: auto; padding: 0 8px 12px; display: flex; flex-direction: column; gap: 1px; }
.history-empty { padding: 16px; font-size: 12px; color: var(--text3); text-align: center; }
.history-row { display: flex; align-items: stretch; border-radius: var(--r); overflow: hidden; }
.history-row:hover .history-btn { background: var(--surface2); }
.history-row:hover .history-del { opacity: 1; }
.history-btn { flex: 1; min-width: 0; text-align: left; padding: 8px 10px; border: none; background: transparent; cursor: pointer; }
.history-btn.active { background: var(--accent-bg); }
.history-name { font-size: 12px; font-weight: 500; color: var(--text); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.history-meta { font-size: 11px; color: var(--text3); margin-top: 1px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.history-del { opacity: 0; background: none; border: none; color: var(--text3); font-size: 13px; padding: 0 8px; transition: opacity .15s, color .15s; flex-shrink: 0; }
.history-del:hover { color: var(--danger); }
.av-row { margin-left: 10px; border-left: 2px solid var(--accent-bg); }
.av-tag { display: inline-flex; align-items: center; justify-content: center; width: 14px; height: 14px; border-radius: 3px; background: var(--accent); color: #fff; font-size: 9px; font-weight: 700; margin-right: 5px; flex-shrink: 0; }
.sidebar-footer { padding: 8px 12px 14px; border-top: 0.5px solid var(--border); }
.btn-clear { width: 100%; padding: 7px 10px; border-radius: var(--r); border: 0.5px solid #F09595; background: transparent; color: var(--danger); font-size: 12px; font-weight: 500; transition: background .12s; }
.btn-clear:hover { background: var(--danger-bg); }

.main { flex: 1; display: flex; flex-direction: column; min-width: 0; background: var(--bg); }
.result-topbar { background: var(--surface); border-bottom: 0.5px solid var(--border); padding: 18px 32px 16px; flex-shrink: 0; }
.result-tag { display: inline-flex; align-items: center; gap: 5px; font-size: 11px; font-weight: 500; color: var(--accent); background: var(--accent-bg); padding: 3px 8px; border-radius: 4px; margin-bottom: 6px; }
.result-title { font-size: 20px; font-weight: 500; color: var(--text); margin-bottom: 3px; }
.result-sub { font-size: 13px; color: var(--text2); }
.result-actions { display: flex; gap: 8px; margin-top: 14px; align-items: center; flex-wrap: wrap; }
.btn { display: inline-flex; align-items: center; gap: 6px; padding: 7px 13px; border-radius: var(--r); border: 0.5px solid var(--border2); background: var(--surface); color: var(--text); font-weight: 500; font-size: 13px; transition: background .12s; white-space: nowrap; }
.btn:hover { background: var(--surface2); }
.btn.primary { background: var(--accent); color: #fff; border-color: var(--accent); }
.btn.primary:hover { background: var(--accent-dark); }
.btn.back { color: var(--text2); }

.content { flex: 1; overflow-y: auto; padding: 28px 32px; }
.extract-wrap { max-width: 680px; }
.drop-zone { border: 1.5px dashed var(--border2); border-radius: var(--rl); padding: 52px 24px; text-align: center; cursor: pointer; transition: border-color .2s, background .2s; background: var(--surface); }
.drop-zone:hover, .drop-zone.dragging { border-color: var(--accent); background: var(--accent-bg); }
.drop-zone.disabled { opacity: .5; cursor: not-allowed; pointer-events: none; }
.drop-icon { color: var(--text3); margin-bottom: 14px; }
.drop-title { font-size: 15px; font-weight: 500; color: var(--text); margin-bottom: 4px; }
.drop-sub { font-size: 13px; color: var(--text3); }
.file-queue { display: flex; flex-direction: column; gap: 6px; margin-top: 14px; }
.queue-item { display: flex; align-items: center; gap: 8px; padding: 9px 14px; background: var(--surface); border: 0.5px solid var(--border); border-radius: var(--r); font-size: 13px; color: var(--text2); }
.queue-item.loading { border-color: #B5D4F4; background: var(--accent-bg); }
.queue-item.done    { border-color: #9FE1CB; background: var(--success-bg); }
.queue-item.error   { border-color: #F09595; background: var(--danger-bg); }
.queue-name { font-weight: 500; color: var(--text); flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.queue-size { font-size: 12px; color: var(--text3); flex-shrink: 0; }
.queue-status { font-size: 12px; margin-left: auto; flex-shrink: 0; }
.queue-status.ok  { color: var(--success); }
.queue-status.err { color: var(--danger); cursor: help; }
.queue-remove { margin-left: auto; background: none; border: none; color: var(--text3); font-size: 14px; padding: 0 4px; flex-shrink: 0; }
.queue-remove:hover { color: var(--danger); }
.extract-bar { display: flex; gap: 8px; margin-top: 16px; }
.progress-track { height: 2px; background: var(--surface2); border-radius: 1px; overflow: hidden; margin: 14px 0 6px; }
.progress-bar { height: 100%; width: 0; background: var(--accent); }
.progress-bar.active { animation: prog 12s ease-out forwards; }
@keyframes prog { 0%{width:0%} 50%{width:60%} 90%{width:85%} 100%{width:90%} }
.status-msg { font-size: 12px; color: var(--text3); }

.modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.35); display: flex; align-items: center; justify-content: center; z-index: 100; }
.modal { background: var(--surface); border-radius: var(--rl); border: 0.5px solid var(--border); padding: 24px; width: 500px; max-width: 94vw; max-height: 82vh; display: flex; flex-direction: column; gap: 12px; }
.modal-title { font-size: 15px; font-weight: 600; }
.modal-sub { font-size: 13px; color: var(--text2); }
.modal-list { display: flex; flex-direction: column; gap: 4px; overflow-y: auto; max-height: 260px; }
.modal-bail { display: block; width: 100%; text-align: left; padding: 10px 12px; border-radius: var(--r); border: 0.5px solid var(--border); background: var(--surface); cursor: pointer; transition: background .12s; }
.modal-bail:hover { background: var(--surface2); }
.modal-bail.sel { border-color: var(--accent); background: var(--accent-bg); }
.modal-bail-name { font-size: 13px; font-weight: 500; }
.modal-bail-meta { font-size: 11px; color: var(--text3); margin-top: 2px; }
.modal-actions { display: flex; gap: 8px; justify-content: flex-end; padding-top: 4px; }

.result-body { max-width: 880px; }
.av-banner { display: flex; align-items: flex-start; gap: 10px; background: var(--accent-bg); border: 0.5px solid #B5D4F4; border-radius: var(--r); padding: 12px 16px; font-size: 13px; color: var(--accent-dark); margin-bottom: 24px; }
.av-banner-date { margin-left: auto; font-size: 12px; white-space: nowrap; }

.sec { margin-bottom: 32px; }
.sec-hd { display: flex; align-items: center; gap: 8px; margin-bottom: 12px; }
.sec-dot { width: 5px; height: 5px; border-radius: 50%; background: var(--accent); flex-shrink: 0; }
.sec-label { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .08em; color: var(--text3); }

.g2 { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
.g3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px; }
.g4 { display: grid; grid-template-columns: repeat(4,1fr); gap: 8px; }
.full { grid-column: 1 / -1; }

.field { background: var(--surface); border: 0.5px solid var(--border); border-radius: var(--r); padding: 11px 14px; }
.field-lbl { font-size: 11px; color: var(--text3); margin-bottom: 4px; }
.field-val { font-size: 13px; color: var(--text); font-weight: 500; line-height: 1.5; }
.field-val.empty { color: var(--text3); font-weight: 400; font-style: italic; }
.field-val.mono { font-size: 18px; font-weight: 600; color: var(--accent); letter-spacing: -0.02em; }
.field-val.verbose { font-weight: 400; font-size: 12px; color: var(--text2); line-height: 1.6; }

.party-card { background: var(--surface); border: 0.5px solid var(--border); border-radius: var(--r); padding: 13px 16px; }
.party-role { font-size: 11px; color: var(--text3); text-transform: uppercase; letter-spacing: .06em; margin-bottom: 5px; }
.party-name { font-size: 13px; font-weight: 500; color: var(--text); line-height: 1.5; }

.date-strip { display: grid; gap: 8px; margin-bottom: 8px; }
.date-card { background: var(--surface2); border-radius: var(--r); padding: 11px 14px; }
.date-lbl { font-size: 11px; color: var(--text3); margin-bottom: 3px; display: flex; align-items: center; gap: 4px; }
.date-val { font-size: 14px; font-weight: 500; color: var(--text); }
.date-val.break { color: var(--accent); }
.break-tag { font-size: 10px; background: var(--accent-bg); color: var(--accent); padding: 1px 5px; border-radius: 3px; font-weight: 600; }

.loyer-hero { background: var(--surface2); border-radius: var(--rl); padding: 16px 20px; display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; flex-wrap: wrap; gap: 12px; }
.loyer-amount { font-size: 24px; font-weight: 600; color: var(--text); letter-spacing: -0.02em; }
.loyer-lbl { font-size: 12px; color: var(--text2); margin-bottom: 2px; }
.pills { display: flex; gap: 6px; flex-wrap: wrap; }
.pill { display: inline-flex; align-items: center; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 500; }
.pill-blue  { background: var(--accent-bg); color: var(--accent); }
.pill-green { background: var(--success-bg); color: var(--success); }
.pill-amber { background: var(--amber-bg); color: var(--amber); }

.indemnites-table { width: 100%; border-collapse: collapse; font-size: 12px; }
.indemnites-table th { text-align: left; padding: 7px 10px; font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: .06em; color: var(--text3); border-bottom: 0.5px solid var(--border2); background: var(--surface2); }
.indemnites-table td { padding: 8px 10px; border-bottom: 0.5px solid var(--border); color: var(--text); vertical-align: top; line-height: 1.4; }
.indemnites-table tr:last-child td { border-bottom: none; }
.due-par { display: inline-flex; align-items: center; padding: 1px 7px; border-radius: 999px; font-size: 11px; font-weight: 500; }
.due-preneur { background: var(--danger-bg); color: var(--danger); }
.due-bailleur { background: var(--success-bg); color: var(--success); }

.pair-block { display: grid; grid-template-columns: auto 1fr; gap: 8px; }
.pair-key { background: var(--surface); border: 0.5px solid var(--border); border-radius: var(--r); padding: 11px 14px; min-width: 140px; }
.pair-verbose { background: var(--surface); border: 0.5px solid var(--border); border-left: 2px solid var(--accent-bg); border-radius: var(--r); padding: 11px 14px; }
`

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

  const breaks = Array.isArray(data.break_options) ? data.break_options : data.break_options ? [data.break_options] : []
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

async function callClaude(base64, mediaType, prompt) {
  const res = await fetch('https://vmtmwsbebzkwxfkdpqky.supabase.co/functions/v1/hyper-action', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: 'claude-opus-4-5', max_tokens: 4096,
      messages: [{ role: 'user', content: [
        { type: 'document', source: { type: 'base64', media_type: mediaType, data: base64 } },
        { type: 'text', text: prompt }
      ]}]
    })
  })
  if (!res.ok) throw new Error('Erreur API: ' + res.status)
  const data = await res.json()
  const raw = data.content.map(b => b.text ?? '').join('').trim().replace(/```json|```/g,'').trim()
  const s = raw.indexOf('{'), e = raw.lastIndexOf('}')
  if (s === -1) throw new Error('Pas de JSON')
  return JSON.parse(raw.slice(s, e+1))
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

function IndemniteTable({ indemnites }) {
  if (!Array.isArray(indemnites) || !indemnites.length) return null
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
          {indemnites.map((row, i) => (
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
  const d = isAv ? item.data?.champs_modifies || {} : item.data || {}
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
            {show('loyer_cours') && d.loyer_cours && <Field label="Loyer en cours" value={d.loyer_cours} />}
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

  async function extractOne(file, index) {
    setStatus(index, 'loading')
    const base64 = await toBase64(file)
    const mediaType = getMediaType(file)
    const detectPrompt = `Ce document est-il un bail original ou un avenant ? Réponds UNIQUEMENT : {"type": "bail"} ou {"type": "avenant"}`
    let docType = 'bail'
    try { const r = await callClaude(base64, mediaType, detectPrompt); docType = r?.type === 'avenant' ? 'avenant' : 'bail' } catch(_) {}
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

  function handleClear() { setFiles([]); setStatuses([]); setActiveItem(null) }

  const d = activeItem?.data || {}
  const resultTitle = d.immeuble || d.adresse || activeItem?.file_name || ''
  const resultSub = [d.preneur?.split(',')[0], d.bailleur?.split(',')[0], d.date_signature ? `Signé le ${d.date_signature}` : null].filter(Boolean).join(' · ')

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
                        {!st.state && <button className="queue-remove" onClick={() => { setFiles(p=>p.filter((_,j)=>j!==i)); setStatuses(p=>p.filter((_,j)=>j!==i)) }}>✕</button>}
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
