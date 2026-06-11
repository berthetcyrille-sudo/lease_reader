export const SECTIONS = [
  {
    id: 'actif',
    label: 'Actif',
    fields: [
      { key: 'adresse',   label: 'Adresse' },
      { key: 'immeuble',  label: "Nom de l'immeuble" },
      { key: 'ville',     label: 'Ville / Code postal' },
    ]
  },
  {
    id: 'contrat',
    label: 'Contrat et durée',
    fields: [
      { key: 'type_bail',    label: 'Type de contrat' },
      { key: 'duree_totale', label: 'Durée totale' },
      { key: 'duree_ferme',  label: 'Durée ferme' },
    ]
  },
  {
    id: 'parties',
    label: 'Parties',
    fields: [
      { key: 'preneur',  label: 'Preneur' },
      { key: 'bailleur', label: 'Bailleur' },
      { key: 'garant',   label: 'Garant / Caution' },
    ]
  },
  {
    id: 'dates',
    label: 'Dates clés',
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
      { key: 'loyer_signature',        label: 'Loyer HT/HC à la signature (€)' },
      { key: 'loyer_cours',            label: 'Loyer HT/HC en cours (€)' },
      { key: 'indexation',             label: 'Indexation / indice' },
      { key: 'franchise',              label: 'Franchise' },
      { key: 'participation_travaux',  label: 'Participation travaux bailleur (€)' },
      { key: 'travaux_bailleur_preneur', label: 'Détail travaux financés par le bailleur' },
      { key: 'indemnite_depart',       label: 'Indemnités en cas de départ du preneur' },
      { key: 'depot_garantie',         label: 'Dépôt de garantie' },
      { key: 'charges',                label: 'Charges / TEOM' },
    ]
  },
  {
    id: 'jouissance',
    label: 'Refacturation et jouissance',
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

export const AVENANT_PROMPT = `Tu es un expert en baux commerciaux français. Ce document est un AVENANT à un bail existant.

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
    "adresse": null,
    "immeuble": null,
    "ville": null,
    "type_bail": null,
    "duree_totale": null,
    "duree_ferme": null,
    "preneur": null,
    "bailleur": null,
    "garant": null,
    "date_effet": null,
    "date_signature": null,
    "break_option": null,
    "notice": null,
    "date_conge": null,
    "date_fin": null,
    "date_limite_travaux": null,
    "conditions_break": null,
    "surface_bureaux": null,
    "surface_totale": null,
    "parking": null,
    "rie": null,
    "autres_surfaces": null,
    "loyer_signature": null,
    "loyer_cours": null,
    "indexation": null,
    "franchise": null,
    "participation_travaux": null,
    "travaux_bailleur_preneur": null,
    "indemnite_depart": null,
    "depot_garantie": null,
    "charges": null,
    "article_606": null,
    "conformite": null,
    "accession": null,
    "remise_en_etat": null,
    "maintenance": null,
    "destination": null,
    "sous_location": null,
    "cession": null
  }
}

IMPORTANT : dans "champs_modifies", ne renseigne QUE les champs effectivement modifiés par cet avenant. Laisse null tous les autres. Reprends les montants, dates et formulations exactes du document.`
