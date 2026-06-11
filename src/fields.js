export const SECTIONS = [
  {
    id: 'actif',
    label: 'Actif',
    fields: [
      { key: 'adresse',   label: 'Adresse' },
      { key: 'immeuble',  label: 'Nom de l\'immeuble' },
      { key: 'ville',     label: 'Ville / Code postal' },
    ]
  },
  {
    id: 'contrat',
    label: 'Contrat et durée',
    fields: [
      { key: 'type_bail',     label: 'Type de contrat' },
      { key: 'duree_totale',  label: 'Durée totale' },
      { key: 'duree_ferme',   label: 'Durée ferme' },
    ]
  },
  {
    id: 'parties',
    label: 'Parties',
    fields: [
      { key: 'preneur', label: 'Preneur' },
      { key: 'bailleur', label: 'Bailleur' },
      { key: 'garant',  label: 'Garant / Caution' },
    ]
  },
  {
    id: 'dates',
    label: 'Dates clés',
    fields: [
      { key: 'date_effet',     label: 'Date d\'effet' },
      { key: 'date_signature', label: 'Date de signature' },
      { key: 'break_option',   label: 'Break option' },
      { key: 'notice',         label: 'Préavis (notice)' },
      { key: 'date_conge',     label: 'Date limite de congé' },
      { key: 'date_fin',       label: 'Date de fin' },
    ]
  },
  {
    id: 'surfaces',
    label: 'Surfaces',
    fields: [
      { key: 'surface_bureaux',  label: 'Surface bureaux (m²)' },
      { key: 'surface_totale',   label: 'Surface totale (m²)' },
      { key: 'parking',          label: 'Parking' },
      { key: 'rie',              label: 'RIE' },
      { key: 'autres_surfaces',  label: 'Autres surfaces' },
    ]
  },
  {
    id: 'loyer',
    label: 'Loyer, taxes et charges',
    fields: [
      { key: 'loyer_signature',       label: 'Loyer HT/HC à la signature (€)' },
      { key: 'loyer_cours',           label: 'Loyer HT/HC en cours (€)' },
      { key: 'indexation',            label: 'Indexation / indice' },
      { key: 'franchise',             label: 'Franchise' },
      { key: 'participation_travaux', label: 'Participation travaux bailleur (€)' },
      { key: 'depot_garantie',        label: 'Dépôt de garantie' },
      { key: 'charges',               label: 'Charges / TEOM' },
    ]
  },
  {
    id: 'jouissance',
    label: 'Refacturation et jouissance',
    fields: [
      { key: 'article_606',   label: 'Article 606' },
      { key: 'conformite',    label: 'Conformité' },
      { key: 'accession',     label: 'Accession' },
      { key: 'remise_en_etat', label: 'Remise en état' },
      { key: 'maintenance',   label: 'Maintenance & travaux' },
      { key: 'destination',   label: 'Destination' },
      { key: 'sous_location', label: 'Sous-location' },
      { key: 'cession',       label: 'Cession' },
    ]
  }
]

export const ALL_FIELDS = SECTIONS.flatMap(s => s.fields)

export const EXTRACTION_PROMPT = `Tu es un expert en baux commerciaux français. Analyse ce document et extrait précisément les données suivantes. Réponds UNIQUEMENT en JSON valide, sans markdown ni backticks, avec exactement ces clés :

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
  "surface_bureaux": "détail des surfaces bureaux par niveau/bâtiment",
  "surface_totale": "surface totale en m²",
  "parking": "nombre et description des places de parking",
  "rie": "restaurant inter-entreprises : oui/non et modalités",
  "autres_surfaces": "autres surfaces (archives, locaux techniques, etc.)",
  "loyer_signature": "loyer annuel HT/HC à la signature",
  "loyer_cours": "loyer annuel HT/HC actuel",
  "indexation": "clause d'indexation et indice (ILC, ILAT, ICC...)",
  "franchise": "franchise de loyer (durée, modalités)",
  "participation_travaux": "participation travaux bailleur",
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
