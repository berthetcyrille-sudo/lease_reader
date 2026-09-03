import React, { useState, useEffect, useCallback, useRef, useMemo } from 'react'
import { createPortal } from 'react-dom'
import './index.css'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'
import * as pdfjsLib from 'pdfjs-dist'
import { PDFDocument } from 'pdf-lib'

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/build/pdf.worker.min.mjs', import.meta.url).toString()

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

REGLES: Guillemets droits ASCII. Pas de retour a la ligne dans les valeurs. Champs _montant=chiffres bruts sans symbole (ex: 123405.50). null si absent. INTERDIT: ne JAMAIS concatener une annotation, precision ou commentaire entre parentheses dans un champ date (format strict JJ/MM/AAAA, rien d'autre) ou dans un champ duree (ex: "1 an", "9 ans" — rien d'autre). Si une information complementaire existe (ex: reconduction, plafond de duree, condition), elle DOIT aller dans son champ dedie (ex: reconduction_tacite) et nulle part ailleurs — jamais annexee en texte libre dans date_fin, date_effet ou duree_totale. Exemple INTERDIT: date_fin="31/12/2023 (renouvelable, terme absolu 31/12/2034)" — la valeur correcte est date_fin="31/12/2023" avec reconduction_tacite.date_limite_absolue="31/12/2034" ; duree_totale="1 an renouvelable par tacite reconduction, duree maximale 12 ans" est egalement INTERDIT — la valeur correcte est duree_totale="1 an".

CHAMPS:
{"adresse":null,"immeuble":null,"ville":null,"type_bail":null,"duree_totale":null,"duree_ferme":null,"preneur":null,"bailleur":null,"garant":null,"date_effet":null,"date_signature":null,"break_options":[],"notice":null,"date_conge":null,"date_fin":null,"date_limite_travaux":null,"conditions_break":null,"reconduction_tacite":null,"surface_totale_m2":null,"surfaces_detail":[],"parking_nb_places":null,"parking":null,"rie":null,"loyer_signature_montant":null,"loyer_signature":null,"loyer_cours":null,"indexation":null,"indexation_indice":null,"indexation_trimestre_base":null,"indexation_valeur_base":null,"franchise_periodes":[],"franchise":null,"charges":null,"depot_garantie_montant":null,"depot_garantie":null,"travaux_montant":null,"travaux_date_factures":null,"travaux_modalites":null,"participations_travaux":[],"indemnites":[],"indemnites_detail":null,"article_606":null,"conformite":null,"accession":null,"remise_en_etat":null,"maintenance":null,"destination":null,"sous_location":null,"cession":null,"mise_a_disposition":null,"indemnites_restitution":[],"_sources":{},"_pages":{}}

REGLES PAR CHAMP:
- duree_totale: duree totale du bail (date_effet a date_fin). duree_ferme: duree pendant laquelle le preneur ne peut pas resilier; si mentionne explicitement utiliser cette valeur; si break_options, c'est l'intervalle date_effet->premiere break. IMPORTANT: si duree_ferme < duree_totale et break_options est vide, ajouter dans break_options la date correspondant a date_effet + duree_ferme (premiere sortie possible). ATTENTION: NE JAMAIS mettre duree_ferme = duree_totale par defaut quand rien n'est explicitement restreint — un bail SANS renonciation ni restriction du droit de resiliation triennale (art. L.145-4) a en realite une duree_ferme implicite de 3 ans (premiere sortie possible), PAS une duree_ferme egale a la duree totale (ce qui reviendrait a interdire toute sortie anticipee, ce qui n'est pas ce que dit le bail dans ce cas). Si aucune duree ferme n'est explicitement chiffree ET qu'aucune renonciation totale n'est exprimee, laisser duree_ferme a null plutot que de la deviner egale a duree_totale.
  ATTENTION DEROGATION PARTIELLE CP/CG: quand une clause CP dit "par derogation a l'article CG-X - DUREE, le Contrat est consenti pour une DUREE FERME de N annees, le PRENEUR renoncant a donner conge a l'expiration de la 3eme et de la 6eme annee" (ou formulation equivalente), cette clause CP ne deroge QUE sur le mecanisme de sortie anticipee (duree_ferme = N annees) — elle NE redefinit PAS la duree totale du bail. La duree_totale reste celle fixee par l'article CG-X vise (souvent PLUS LONGUE, ex: 12 ans), sauf si la clause CP dit explicitement autre chose sur la duree totale elle-meme. NE JAMAIS recopier la valeur de la "duree ferme" CP dans duree_totale sans avoir verifie la duree totale dans l'article CG correspondant. Exemple: CP dit "par derogation a l'article CG3 - DUREE, le Contrat est consenti pour une duree ferme de neuf (9) annees, le PRENEUR renoncant a la 3eme et 6eme annee" et CG3 dit "le Contrat est consenti pour une duree de douze (12) annees" → duree_totale=12 ans, duree_ferme=9 ans (PAS duree_totale=9 ans).
- reconduction_tacite: si le bail prevoit qu'au-dela du terme (date_fin), le contrat se poursuit automatiquement par tacite reconduction (annee par annee ou periode similaire) jusqu'a ce qu'une partie donne conge avec un preavis. Format: {"applicable":true,"preavis":"6 mois","periodicite":"annuelle","date_limite_absolue":null}. IMPORTANT: dans ce cas, date_fin reste la date de fin du terme FERME initial (ex: fin de la 9eme annee) — NE PAS la traiter comme une fin definitive du bail, la tacite reconduction est un etat DISTINCT et POSTERIEUR qui se rajoute. date_limite_absolue: certains baux plafonnent la duree totale possible de la reconduction tacite par une clause du type "en tout etat de cause, la Convention/le Contrat ne pourra exceder N (en toutes lettres) annees et prendra automatiquement fin le [date], sans formalite". Si une telle clause EXPLICITE existe, reporter cette date exacte (format JJ/MM/AAAA) dans date_limite_absolue — c'est un plafond contractuel dur, distinct du preavis de conge habituel. Sinon (reconduction tacite sans limite de duree totale exprimee), laisser date_limite_absolue a null. null pour le champ reconduction_tacite entier si le bail prevoit un terme ferme sans reconduction automatique (bail qui s'eteint purement et simplement a date_fin).
- surface_totale_m2: la surface de reference du bail. REGLE: si le bail utilise le terme "Surface Exploitee" (ou variante proche) pour designer la surface globale des locaux, UTILISER CETTE VALEUR pour surface_totale_m2, meme si elle inclut une quote-part des parties communes — c'est la convention de reference dans ce bail. Ne descendre au sous-composant individuel (ex: "Surface de bureaux") QUE si aucune "Surface Exploitee"/surface globale n'est mentionnee. Exemple: "la Surface Exploitee... est de 584,50 m²... les Locaux se decomposent: Surface de bureaux (lot n°11): 510,20 m²" → surface_totale_m2 = 584.50 (la Surface Exploitee), PAS 510.20.
- surfaces_detail: TOUTES les surfaces explicitement chiffrees dans le bail, meme celles sans ventilation de loyer propre. REGLE PRIORITAIRE: des qu'une surface est donnee avec un chiffre (ex: "Surface interieure: 2503 m2", "Surface exterieure/terrasse: 630 m2"), creer une LIGNE DISTINCTE pour elle dans surfaces_detail, MEME SI aucun loyer_annuel specifique n'est indique pour cette surface — dans ce cas mettre loyer_annuel a null pour cette ligne plutot que d'omettre la ligne. NE JAMAIS repartir/dupliquer artificiellement le loyer total (loyer_signature_montant) sur plusieurs lignes quand le bail ne le ventile pas explicitement par composante — laisser loyer_annuel a null sur les lignes non ventilees. Inclure AUSSI les redevances forfaitaires liees a l'usage des surfaces (RIE/restauration, archives, locaux techniques) meme si exprimees en €/m²/an. Exemple avec ventilation de loyer (toutes les lignes ont un loyer_annuel): [{\"categorie\":\"Bureaux\",\"niveau\":\"2eme etage\",\"surface_m2\":\"245.68\",\"prix_unitaire\":\"196\",\"loyer_annuel\":\"48122\"},{\"categorie\":\"RIE\",\"niveau\":\"RDC\",\"surface_m2\":\"245.68\",\"prix_unitaire\":\"15\",\"loyer_annuel\":\"3685\"}]. Exemple SANS ventilation de loyer par composante (loyer global uniquement): bail dit "Surface interieure: 2503 m2, Surface exterieure: 630 m2" et "redevance annuelle: 362935 EUR HT" sans repartition → [{\"categorie\":\"Bureaux\",\"niveau\":\"1er etage - interieur\",\"surface_m2\":\"2503\",\"loyer_annuel\":null},{\"categorie\":\"Terrasse\",\"niveau\":\"1er etage - exterieur\",\"surface_m2\":\"630\",\"loyer_annuel\":null}] (loyer_signature_montant=362935 reste renseigne separement, PAS reparti sur ces 2 lignes). categorie: etage/plateau->Bureaux, terrasse/rooftop/exterieur->Terrasse, sous-sol/emplacement->Stationnement, restaurant/cafeteria/restauration->RIE (Restaurant Inter-Entreprises), archives->Archives, reserves/stockage->Archives. Si TOUTES les lignes ont un loyer_annuel renseigne, leur SOMME doit etre egale a loyer_signature_montant — cette regle ne s'applique PAS quand une ou plusieurs lignes ont loyer_annuel=null (pas de ventilation disponible). Si le bail mentionne une "Surface Exploitee" distincte des sous-composantes louees (incluant une quote-part de parties communes), la somme des surface_m2 peut legitimement etre INFERIEURE a surface_totale_m2 — ce n'est pas une erreur a corriger dans ce cas.
- notice: DUREE du préavis pour donner congé, exprimée en mois uniquement (ex: "6 mois", "3 mois"). NE PAS mettre une date. Si le bail dit "au moins six (6) mois avant la date d'échéance" → notice="6 mois".
- _sources: objet optionnel avec les extraits textuels EXACTS du bail pour les champs importants. Format: {"loyer_signature_montant":"texte exact de la clause loyer","break_options":"texte exact de la clause duree/resiliation","duree_ferme":"texte exact","franchise_periodes":"texte exact"}. Citer le numero d'article si possible (ex: "CP4 - Le loyer annuel est de..."). Limiter a 150 caracteres par champ.
- _pages: objet avec le numero de PAGE du PDF (1=premiere page) ou se trouve l'information source, pour chaque champ dont la valeur n'est pas null. Format: {"loyer_signature_montant":3,"date_effet":1,"date_fin":1,"break_options":4,"duree_totale":1,"duree_ferme":1,"surface_totale_m2":2,"preneur":1,"bailleur":1,"depot_garantie_montant":5}. Indiquer la page pour un maximum de champs renseignes (duree_totale et duree_ferme sont presque toujours dans la meme clause, ne pas en oublier un des deux), meme approximative si le champ resulte d'un calcul (prendre la page de la clause source utilisee pour le calcul). Ne pas inclure les champs restes null.
- mise_a_disposition: si le bail prevoit une mise a disposition anticipee des locaux (avant la date d'effet officielle du bail). Format: {"date_debut":"jj/mm/aaaa","date_fin":"jj/mm/aaaa","loyer_paye":"Oui/Non/Partiel","charges_payees":"Oui/Non/Partiel","conditions":"texte libre des conditions financieres pendant cette periode"}. null si aucune mise a disposition anticipee.
- break_options: liste COMPLETE et EXHAUSTIVE de toutes les dates auxquelles le PRENEUR peut effectivement sortir avant le terme. Format: ["31/08/2028","31/08/2031"]. REGLE CRITIQUE: les CP priment TOUJOURS sur les CG — SAUF si les CP renvoient EXPRESSEMENT a un article des CG pour les modalites de conge/duree (ex: "le PRENEUR pourra delivrer conge selon les modalites convenues a l'article CG2.2 DUREE CONGE du Bail") : dans ce cas, c'est CET ARTICLE DES CG qu'il faut lire et appliquer, pas l'ignorer sous pretexte que les CP priment en general — les CP eux-memes designent les CG comme source de la regle. REGLE ABSOLUE CONVENTION DE DATE: toute date calculee par addition d'un nombre entier d'annees a date_effet (breaks, ou date_fin si elle doit etre calculee) s'exprime au JOUR ANNIVERSAIRE MOINS 1 JOUR — convention standard des baux commerciaux francais. Exemple: date_effet=15/10/2020, echeance a 6 ans → 15/10/2020 + 6 ans = 15/10/2026, MOINS 1 JOUR = 14/10/2026 (PAS 15/10/2026). Cette regle s'applique a TOUTE date calculee dans break_options ET a date_fin lorsqu'elle doit etre deduite de duree_totale (mais PAS si le bail donne une date de fin EXPLICITE en toutes lettres — dans ce cas utiliser cette date telle quelle, meme si elle ne suit pas cette convention). REGLE ABSOLUE (s'applique a TOUTES les regles ci-dessous, quelle que soit la formulation du bail: "periode triennale", "echeance triennale", "faculte triennale", etc.): les echeances triennales SUCCESSIVES se calculent TOUJOURS comme des multiples de 3 ANS DEPUIS DATE_EFFET (annees 3, 6, 9, 12...), JAMAIS en ajoutant 3 ans a la date de la break precedente. Meme si le bail nomme une premiere sortie a une annee qui n'est pas un multiple de 3 (ex: annee 4, ou annee 7, en fin de duree ferme), l'echeance suivante reste le PROCHAIN multiple de 3 depuis date_effet strictement superieur a cette annee — PAS cette annee + 3. REGLES DE CALCUL:
  1) "a l'expiration de chaque periode triennale" → date_effet + 3 ans, + 6 ans, + 9 ans (si < date_fin)
  2) "renonce a sa faculte de resiliation triennale" SANS restriction → aucune break triennale, uniquement les dates explicites des CP
  3) "renonce...triennale POUR LA DUREE FERME" ou "aura la faculte de donner conge a l'expiration de la Neme periode/echeance triennale pour la premiere fois" → premiere break = date_effet + N*3 ans. PUIS CONTINUER a intervalles de 3 ans supplementaires (N+1, N+2...) TANT QUE la date obtenue reste STRICTEMENT ANTERIEURE a date_fin — ne JAMAIS s'arreter apres la premiere date par defaut. Exemple A (plusieurs breaks): "deuxieme periode triennale pour la premiere fois", date_effet=01/01/2021, date_fin=31/12/2030 → premiere break=31/12/2026 (2*3=6 ans) ; echeance suivante=31/12/2029 (3*3=9 ans), qui est < date_fin donc AJOUTEE aussi → ["31/12/2026","31/12/2029"]. Exemple B (un seul break, car l'echeance suivante coincide avec la fin du bail): meme formulation, date_effet=30/06/2025, date_fin=29/06/2034 → premiere break=30/06/2031 (6 ans) ; echeance suivante=30/06/2034 (9 ans) qui EGALE (a 1 jour pres) date_fin donc EXCLUE → ["30/06/2031"] uniquement
  3bis) "renonce a sa faculte de resiliation triennale pour la duree ferme" avec duree_ferme exprimee en annees NON multiple de 3 (ex: 4 ou 7 ans) et le texte precise une premiere sortie explicite ("au plus tot le [date]", "a l'expiration de la Neme annee...pour la premiere fois", "soit le [date] pour la premiere fois") suivie ou non d'une reference a une echeance/periode triennale suivante SANS date explicite ("ainsi que pour la deuxieme echeance triennale", "et a chaque echeance triennale suivante") → cette premiere date (= date_effet + duree_ferme, meme si duree_ferme n'est pas un multiple de 3) est la premiere break. La ou les echeance(s) suivante(s), MEME NOMMEES "deuxieme echeance" dans le texte, se calculent TOUJOURS depuis date_effet (PROCHAIN multiple de 3 strictement superieur a duree_ferme), JAMAIS depuis la premiere break. EXEMPLE 1 (duree_ferme=7 ans): date_effet=01/09/2022, "conge au plus tot pour le 31 aout 2029" → premiere break=31/08/2029 (annee 7) ; prochain multiple de 3 apres 7 = annee 9 → echeance suivante=31/08/2031 (PAS annee 7+3=10=2032) ; date_fin=31/08/2034 (annee 12, EXCLUE) → ["31/08/2029","31/08/2031"]. EXEMPLE 2 (duree_ferme=4 ans): date_effet=01/05/2024, "conge a l'expiration de la quatrieme annee, soit le 30 avril 2028 pour la premiere fois, ainsi que pour la deuxieme echeance triennale" → premiere break=30/04/2028 (annee 4) ; prochain multiple de 3 apres 4 = annee 6 (PAS annee 4+3=7=2031, meme si le texte dit "deuxieme echeance" juste apres la premiere) → echeance suivante=30/04/2030 ; date_fin=30/04/2033 (annee 9, EXCLUE) → ["30/04/2028","30/04/2030"]
  4) "a l'expiration de la Neme annee" → date_effet + N ans
  5) "renoncant expressement a sa faculte de donner conge a l'expiration de la Neme et de la Mieme Annee du Contrat" (ENUMERATION EXPLICITE des annees renoncees, formulation DIFFERENTE de "duree ferme" ou "Nieme fois") → calculer TOUTES les echeances triennales standard (multiples de 3 depuis date_effet: annee 3, 6, 9, 12...) jusqu'a duree_totale, PUIS RETIRER les annees explicitement nommees comme renoncees. NE JAMAIS inclure une annee que le texte nomme explicitement comme renoncee, meme si elle tombe sur un multiple de 3 par ailleurs. EXEMPLE (celui-ci est un piège frequent — l'annee renoncee est un vrai multiple de 3, ne pas l'inclure par erreur): date_effet=31/03/2021, duree_totale=12 ans (date_fin=30/03/2033), "renoncant a sa faculte de donner conge a l'expiration de la 3eme et de la 6eme Annee" → echeances triennales standard = annees 3,6,9,12 → retirer les annees 3 ET 6 (explicitement renoncees, PAS seulement la 3eme) → il ne reste QUE l'annee 9 (annee 12 de toute facon exclue car = date_fin) → break_options=["30/03/2030"] UNIQUEMENT (PAS "30/03/2027", qui correspond a l'annee 6 explicitement renoncee)
  Ne PAS inclure date_fin.
- loyer_signature_montant: MONTANT ANNUEL TOTAL HT/HC. JAMAIS prix unitaire/m². Si tableau par lot: additionner les loyer_annuel. INTERDIT de retourner null si un loyer figure dans le document.
- loyer_cours: loyer annuel "de base" au sens indexation. Identique a loyer_signature_montant sauf mention contraire. JAMAIS prix unitaire/m².
- indexation_indice: code de l indice parmi: "ILAT","ILC","ICC","IRL","IPC","BT01","AUTRE". null si non renseigne.
- indexation_trimestre_base: trimestre de reference si EXPLICITEMENT indique dans le bail (ex: "3T2025"). null si le bail dit "dernier indice publie a la date de signature" sans preciser lequel.
- indexation_valeur_base: valeur numerique de l indice si EXPLICITEMENT mentionnee (ex: "120.5"). null si non mentionnee.
- franchise_periodes: TOUTES les franchises, y compris conditionnelles. [{\"date_debut\":\"jj/mm/aaaa\",\"date_fin\":\"jj/mm/aaaa\",\"duree\":\"6 mois\",\"montant\":\"123405\",\"surface_assiette\":\"LC1 (701 m²)\",\"indexation_incluse\":\"Non\",\"condition\":null}]. montant=chiffres bruts (calcule si non explicite: loyer_annuel_assiette*duree_mois/12). condition=texte si conditionnelle, null sinon.
- participations_travaux: UNIQUEMENT si le bail prevoit une enveloppe financiere DISTINCTE de la franchise, specifiquement dediee aux travaux (ex: "le BAILLEUR verse X euros pour les travaux" avec un calendrier de facturation propre). EXCLURE: les franchises de loyer qualifiees de participation aux travaux (ex: "franchise accordee au titre de la participation aux travaux") — ces franchises doivent figurer UNIQUEMENT dans franchise_periodes. En cas de doublon franchise/travaux sur le meme montant, privilegier franchise_periodes. Format: [{\"libelle\":\"denomination exacte\",\"montant\":\"822701\",\"date_limite\":\"31/12/2024\",\"remarque\":null}]. libelle OBLIGATOIRE.
- parking_nb_places: ex: "114 places (98 interieures + 16 exterieures)"
- indemnites: UNIQUEMENT indemnites liees a une option (break, renouvellement, fin de bail). EXCLURE: honoraires, cautionnements, penalites. [{\"motif\":\"...\",\"due_par\":\"Preneur ou Bailleur\",\"montant\":\"chiffres bruts\",\"date_limite\":\"...\"}]`

// Prompt léger pour rattraper les documents déjà extraits avant l'ajout du
// stockage : on NE redemande PAS d'extraire les données (déjà en base, potentiellement
// déjà vérifiées manuellement), seulement de localiser leur page dans le PDF.
function buildLocatePagesPrompt(values) {
  return `Voici des valeurs déjà extraites d'un document (bail ou avenant commercial français). Pour CHAQUE valeur ci-dessous qui apparaît clairement dans ce document, indique le numéro de PAGE du PDF (1 = première page) où elle se trouve. N'invente rien : si tu ne retrouves pas une valeur avec confiance, omets-la simplement. Réponds UNIQUEMENT avec un JSON minifié sur une seule ligne, sans markdown, format : {"champ1":page,"champ2":page}.

Valeurs à localiser :
${JSON.stringify(values)}`
}

const AVENANT_PROMPT = `Expert baux commerciaux français. Ce document est un AVENANT. JSON minifié UNE SEULE LIGNE, sans markdown.

REGLES: Guillemets droits ASCII. Champs _montant=chiffres bruts. Dans champs_modifies: null pour les champs NON modifies par l'avenant.

surface_change_type: "inchangee"/"ajout"/"retrait"/"substitution"/"mixte".
surfaces_delta: surfaces UNIQUEMENT concernees par la modif (ajoutees ou retirees). Ajouter "sens":"ajout" ou "sens":"retrait". categorie JAMAIS null.
surfaces_avant: tableau EXACT des surfaces telles qu'elles etaient AVANT cet avenant, tel que decrit dans le bail d'origine mentionne dans ce document. categorie JAMAIS null. null si surface_change_type="inchangee".
surfaces_apres: tableau EXACT des surfaces APRES cet avenant. REGLE STRICTE: regrouper par categorie si plusieurs lignes de meme categorie (ex: 2 lignes Bureaux → une seule ligne avec la surface totale). NE PAS INVENTER de lignes. NE PAS dupliquer. La surface totale de surfaces_apres doit etre egale a surface_totale_m2. categorie JAMAIS null. null si surface_change_type="inchangee".

{"bail_reference":{"preneur":null,"bailleur":null,"date_bail_origine":null,"adresse":null,"immeuble":null},"date_effet_avenant":null,"date_signature_avenant":null,"objet_avenant":null,"surface_change_type":"inchangee","surfaces_delta":null,"surfaces_avant":null,"surfaces_apres":null,"champs_modifies":{"adresse":null,"immeuble":null,"ville":null,"type_bail":null,"duree_totale":null,"duree_ferme":null,"preneur":null,"bailleur":null,"garant":null,"date_effet":null,"date_signature":null,"break_options":null,"notice":null,"date_conge":null,"date_fin":null,"date_limite_travaux":null,"conditions_break":null,"reconduction_tacite":null,"surface_totale_m2":null,"surfaces_detail":null,"parking_nb_places":null,"parking":null,"rie":null,"loyer_signature_montant":null,"loyer_signature":null,"loyer_cours":null,"indexation":null,"franchise_periodes":null,"franchise":null,"charges":null,"depot_garantie_montant":null,"depot_garantie":null,"travaux_montant":null,"travaux_date_factures":null,"travaux_modalites":null,"participations_travaux":null,"indemnites":null,"indemnites_detail":null,"article_606":null,"conformite":null,"accession":null,"remise_en_etat":null,"maintenance":null,"destination":null,"sous_location":null,"cession":null,"mise_a_disposition":null,"indemnites_restitution":[],"_sources":{}},"_pages":{}}

REGLES PAR CHAMP (champs_modifies):
- loyer_signature_montant: montant annuel total HT/HC. null si non modifie. JAMAIS prix unitaire/m².
- break_options: UNIQUEMENT si l'avenant modifie/redefinit les dates de sortie anticipee. Format: TABLEAU DE DATES PURES au format "jj/mm/aaaa" UNIQUEMENT, ex: ["31/12/2030","31/12/2033"]. JAMAIS de phrase descriptive (interdit: "Premiere faculte de conge a l'expiration de la 2e periode triennale le 31/12/2030" — mettre uniquement "31/12/2030"). Si l'avenant dit "renonciation a la resiliation triennale pour la duree ferme de N ans" ou "premier conge possible le jj/mm/aaaa", extraire la ou les date(s) exacte(s) mentionnee(s), pas le texte de la clause (le texte de la clause va dans conditions_break et _sources, pas dans break_options). null si non modifie.
- franchise_periodes: TOUTES les nouvelles franchises de l'avenant. [{\"date_debut\":\"jj/mm/aaaa\",\"date_fin\":\"jj/mm/aaaa\",\"duree\":\"6 mois\",\"montant\":\"123405\",\"surface_assiette\":\"LC1 (701 m²)\",\"indexation_incluse\":\"Non\",\"condition\":null}]. null si aucune franchise dans l'avenant.
- participations_travaux: UNIQUEMENT si enveloppe financiere DISTINCTE de la franchise, dediee aux travaux avec calendrier de facturation propre. Ne JAMAIS y mettre une franchise de loyer meme si qualifiee "au titre des travaux" — celle-ci va dans franchise_periodes. En cas de doute sur meme montant, privilegier franchise_periodes. Format: [{\"libelle\":\"denomination exacte\",\"montant\":\"822701\",\"date_limite\":\"31/12/2024\",\"remarque\":null}]. null si non concerne.
- surfaces_detail: tableau complet post-avenant UNIQUEMENT si l'avenant redefinit completement l'assiette. null sinon (utiliser surfaces_apres a la place).
- _pages: objet AU MEME NIVEAU que champs_modifies (pas dedans) avec le numero de PAGE du PDF ou se trouve chaque champ RENSEIGNE (non-null) de champs_modifies, plus objet_avenant, date_effet_avenant et date_signature_avenant si applicable. Format: {"loyer_signature_montant":2,"date_effet_avenant":1,"objet_avenant":1}. Ne pas inclure les champs restes null.`

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

// ─── Stockage du fichier source ─────────────────────────────────────────────
// Upload non-bloquant : si ça échoue, l'extraction reste valide, on perd juste
// la possibilité de revoir le fichier d'origine depuis le dashboard.
function sanitizeStorageKey(name) {
  return name
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // retire les accents (é -> e)
    .replace(/[^a-zA-Z0-9._-]+/g, '_') // remplace espaces et caractères spéciaux
}

async function uploadSourceFile(recordId, file) {
  try {
    const path = `${recordId}/${sanitizeStorageKey(file.name)}`
    const { error: uploadError } = await supabase.storage
      .from('lease-sources')
      .upload(path, file, { upsert: true, contentType: file.type || getMediaType(file) })
    if (uploadError) throw uploadError
    const { error: updateError } = await supabase.from('extractions')
      .update({ storage_path: path }).eq('id', recordId)
    if (updateError) throw updateError
    return path
  } catch (e) {
    console.error('Upload du fichier source échoué pour', file.name, e)
    return null
  }
}

// ─── Ouverture du fichier source à une page précise ─────────────────────────
// Beaucoup de navigateurs (Chrome, Firefox, Edge) honorent le fragment
// #page=N sur une URL PDF ouverte directement — pas besoin de lecteur maison.
async function openSourceAtPage(item, page) {
  if (!item?.storage_path) return
  try {
    const { data, error } = await supabase.storage
      .from('lease-sources')
      .createSignedUrl(item.storage_path, 60)
    if (error) throw error
    window.open(`${data.signedUrl}#page=${page || 1}`, '_blank')
  } catch (e) {
    console.error('Impossible d\'ouvrir le fichier source', e)
  }
}

// Petite icône cliquable renvoyant à la page source d'un champ, si connue.
function PageJumpIcon({ item, pages, field, title }) {
  const page = pages?.[field]
  if (!page || !item?.storage_path) return null
  return (
    <span
      onClick={e => { e.stopPropagation(); openSourceAtPage(item, page) }}
      title={title || `Voir la page ${page} du document source`}
      style={{
        cursor: 'pointer', fontSize: '10px', fontWeight: 700, color: 'var(--accent)',
        border: '1px solid rgba(26,95,168,.3)', borderRadius: '4px', padding: '0px 4px',
        display: 'inline-flex', alignItems: 'center', gap: '2px', whiteSpace: 'nowrap',
        background: 'var(--accent-bg)', lineHeight: '15px',
      }}>
      p.{page}
    </span>
  )
}

function getMediaType(file) {
  return file.name.toLowerCase().endsWith('.pdf')
    ? 'application/pdf'
    : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
}

// ─── Compression PDF automatique ────────────────────────────────────────────
// Certains PDF (ex. "Print to PDF", scans) sont en réalité des empilements de
// calques image très lourds sans texte sélectionnable. Recompresser les images
// une par une casse ces calques (transparence/empilement). La méthode fiable :
// rasteriser chaque page déjà composée par le moteur de rendu, puis reconstruire
// un PDF léger à partir de ces images. Ne s'applique qu'aux PDF dépassant le seuil.

const PDF_COMPRESS_THRESHOLD = 8 * 1024 * 1024 // 8 Mo : en-dessous, on ne touche à rien
const PDF_RENDER_DPI = 150
const PDF_JPEG_QUALITY = 0.72

async function compressPdfIfNeeded(file, onProgress) {
  if (!file.name.toLowerCase().endsWith('.pdf')) return file
  if (file.size <= PDF_COMPRESS_THRESHOLD) return file

  try {
    const arrayBuffer = await file.arrayBuffer()
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise
    const numPages = pdf.numPages
    const outDoc = await PDFDocument.create()
    const scale = PDF_RENDER_DPI / 72 // pdf.js viewport de base = 72dpi

    for (let i = 1; i <= numPages; i++) {
      onProgress?.(i, numPages)
      const page = await pdf.getPage(i)
      const viewport = page.getViewport({ scale })
      const canvas = document.createElement('canvas')
      canvas.width = Math.ceil(viewport.width)
      canvas.height = Math.ceil(viewport.height)
      const ctx = canvas.getContext('2d')
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, canvas.width, canvas.height)
      await page.render({ canvasContext: ctx, viewport }).promise

      const blob = await new Promise(res => canvas.toBlob(res, 'image/jpeg', PDF_JPEG_QUALITY))
      const jpegBytes = new Uint8Array(await blob.arrayBuffer())
      const jpgImage = await outDoc.embedJpg(jpegBytes)
      const pdfPage = outDoc.addPage([viewport.width, viewport.height])
      pdfPage.drawImage(jpgImage, { x: 0, y: 0, width: viewport.width, height: viewport.height })

      canvas.width = 0
      canvas.height = 0
    }

    const outBytes = await outDoc.save()
    const compressed = new File([outBytes], file.name, { type: 'application/pdf' })
    // Garde-fou : si jamais la compression ne suffit pas (cas extrême), on renvoie
    // quand même le résultat compressé, qui sera toujours plus léger que l'original.
    return compressed.size < file.size ? compressed : file
  } catch (e) {
    console.error('Compression PDF échouée pour', file.name, e)
    return file // on retombe sur le fichier original ; le contrôle de taille prendra le relais
  }
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
  let rows, statuts
  if (Array.isArray(items)) {
    rows = items.map(({ item, parentName, parentData }) => buildExcelRow(item, parentName, parentData))
    statuts = items.map(({ statut }) => statut || 'OK')
  } else {
    const fakeItem = { document_type: 'bail', data: items, file_name: fileName }
    rows = [buildExcelRow(fakeItem, '')]
    statuts = ['OK']
  }

  const headers = ['Statut', ...buildExcelHeaders()]
  const dataRows = rows.map((row, i) => [statuts[i], ...row])
  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.aoa_to_sheet([headers, ...dataRows])

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
    const isAmount = h.includes('Montant') || h.includes('montant') || h.includes('Loyer') || h.includes('loyer') || h.includes('Prix')
    const isSurface = h.includes('m2') || h.includes('M2')
    if (!isAmount && !isSurface) return
    dataRows.forEach((_, rowIdx) => {
      const cell = ws[XLSX.utils.encode_cell({ r: rowIdx + 1, c: colIdx })]
      if (cell && typeof cell.v === 'number') {
        cell.t = 'n'
        cell.z = isSurface ? '#,##0.00' : '#,##0.##'
      }
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

function exportAllToExcel(tree, onErrors) {
  const rows = []
  const errors = [] // { name, reason }

  tree.forEach(bail => {
    const parentName = bail.data?.immeuble || bail.data?.adresse || bail.file_name
    const parentData = bail.data || {}
    rows.push({ item: bail, parentName: '', parentData: null, statut: 'OK' })
    const sortedAv = [...(bail.avenants || [])].sort((a, b) => {
      const toS = d => { const m = String(d||'').match(/^(\d{2})\/(\d{2})\/(\d{4})$/); return m ? `${m[3]}-${m[2]}-${m[1]}` : String(d||'') }
      return toS(a.data?.date_effet_avenant || a.data?.date_signature_avenant || a.created_at)
            .localeCompare(toS(b.data?.date_effet_avenant || b.data?.date_signature_avenant || b.created_at))
    })
    sortedAv.forEach(av => rows.push({ item: av, parentName, parentData, statut: 'OK' }))
  })

  // Detect orphan avenants (no parent in tree)
  const bailIds = new Set(tree.map(b => b.id))
  // Also check history for orphan avenants passed separately
  // Orphans are avenants in tree with no parent → they appear as top-level with document_type=avenant
  tree.filter(r => r.document_type === 'avenant').forEach(av => {
    errors.push({
      name: av.data?.immeuble || av.data?.adresse || av.file_name,
      reason: 'Avenant orphelin — bail parent manquant ou en erreur'
    })
    // Add to export with warning status
    rows.push({ item: av, parentName: '', parentData: null, statut: '⚠ Bail parent manquant' })
  })

  exportToExcel(rows, 'lease_abstract_complet')
  if (errors.length > 0) onErrors?.(errors)
}

const BREAK_PROMPT = `Expert baux commerciaux français. Analyse UNIQUEMENT la clause de durée et de résiliation de ce bail. Retourne UNIQUEMENT un JSON minifié sur UNE SEULE LIGNE : {"date_effet":"jj/mm/aaaa","date_fin":"jj/mm/aaaa","break_options":["jj/mm/aaaa",...]}

REGLE ABSOLUE pour break_options : liste COMPLETE et EXHAUSTIVE.
REGLE ABSOLUE CONVENTION DE DATE: toute date calculee par addition d'annees entieres a date_effet s'exprime au JOUR ANNIVERSAIRE MOINS 1 JOUR (convention standard des baux commerciaux). Exemple: date_effet=15/10/2020, echeance a 6 ans → 14/10/2026 (PAS 15/10/2026). Ne s'applique pas si le bail donne une date EXPLICITE en toutes lettres.
REGLE ABSOLUE (quelle que soit la formulation: "periode triennale", "echeance triennale", etc.): les echeances triennales SUCCESSIVES se calculent TOUJOURS comme des multiples de 3 ANS DEPUIS DATE_EFFET, JAMAIS en ajoutant 3 ans a la break precedente — meme si la premiere sortie nommee par le bail tombe sur une annee qui n'est pas un multiple de 3 (ex: annee 4 ou 7, fin de duree ferme).
REGLES DE CALCUL (lire attentivement la clause, ne pas appliquer mecaniquement):
1) "a l'expiration de chaque periode triennale" → date_effet + 3 ans, + 6 ans, + 9 ans (si < date_fin)
2) "renonce expressement a sa faculte de resiliation triennale" SANS restriction → PAS de break triennale, uniquement dates explicites des CP
3) "renonce...triennale POUR LA DUREE FERME" OU "aura la faculte de donner conge a l'expiration de la Neme periode/echeance triennale pour la premiere fois" → premiere break = date_effet + N*3 ans. PUIS CONTINUER a intervalles de 3 ans supplementaires (N+1, N+2...) TANT QUE la date obtenue reste STRICTEMENT ANTERIEURE a date_fin — ne JAMAIS s'arreter apres la premiere date par defaut. EXEMPLE A (plusieurs breaks): "deuxieme periode triennale pour la premiere fois", date_effet=01/01/2021, date_fin=31/12/2030 → premiere break=31/12/2026 (6 ans) ; echeance suivante=31/12/2029 (9 ans) < date_fin donc AJOUTEE → ["31/12/2026","31/12/2029"]. EXEMPLE B (un seul, car l'echeance suivante coincide avec la fin du bail): meme formulation, date_effet=30/06/2025, date_fin=29/06/2034 → premiere break=30/06/2031 (6 ans) ; echeance suivante=30/06/2034 (9 ans) = date_fin (a 1 jour pres) donc EXCLUE → ["30/06/2031"] uniquement
3bis) "renonce a sa faculte de resiliation triennale pour la duree ferme" avec duree_ferme NON multiple de 3 (ex: 4 ou 7 ans) et une premiere sortie explicite ("au plus tot", "pour la premiere fois", "a l'expiration de la Neme annee") EVENTUELLEMENT suivie d'une reference a une echeance triennale suivante SANS date explicite ("ainsi que pour la deuxieme echeance triennale") → cette premiere date (date_effet + duree_ferme) est la premiere break. La ou les echeance(s) suivante(s), MEME NOMMEE "deuxieme echeance" dans le texte juste apres, se calculent TOUJOURS depuis date_effet (PROCHAIN multiple de 3 strictement superieur a duree_ferme), JAMAIS depuis la premiere break. EXEMPLE 1 (duree_ferme=7 ans): date_effet=01/09/2022, conge au plus tot 31/08/2029 → premiere break=31/08/2029 (annee 7) ; prochain multiple de 3 apres 7 = annee 9 → echeance suivante=31/08/2031 (PAS annee 7+3=2032) ; date_fin=31/08/2034 (annee 12, EXCLUE) → ["31/08/2029","31/08/2031"]. EXEMPLE 2 (duree_ferme=4 ans): date_effet=01/05/2024, "conge a l'expiration de la 4eme annee, soit le 30/04/2028 pour la premiere fois, ainsi que pour la deuxieme echeance triennale" → premiere break=30/04/2028 (annee 4) ; prochain multiple de 3 apres 4 = annee 6 (PAS annee 4+3=2031) → echeance suivante=30/04/2030 ; date_fin=30/04/2033 (annee 9, EXCLUE) → ["30/04/2028","30/04/2030"]
4) "a l'expiration de la Neme annee" → date_effet + N ans
5) "renoncant expressement a sa faculte de donner conge a l'expiration de la Neme et de la Mieme Annee du Contrat" (ENUMERATION EXPLICITE des annees renoncees, formulation DIFFERENTE de "duree ferme" ou "Nieme fois") → calculer TOUTES les echeances triennales standard (multiples de 3 depuis date_effet: annee 3,6,9,12...) jusqu'a duree_totale, PUIS RETIRER les annees explicitement nommees comme renoncees — meme si une annee renoncee tombe sur un multiple de 3. EXEMPLE (piege frequent): date_effet=31/03/2021, duree_totale=12 ans (date_fin=30/03/2033), "renoncant a sa faculte de donner conge a l'expiration de la 3eme et de la 6eme Annee" → echeances standard = annees 3,6,9,12 → retirer 3 ET 6 (toutes les deux explicitement renoncees) → il ne reste que l'annee 9 (annee 12 exclue = date_fin) → break_options=["30/03/2030"] UNIQUEMENT (PAS "30/03/2027" = annee 6, explicitement renoncee)
CP priment toujours sur CG. Trier chronologiquement. Ne PAS inclure date_fin.`

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
A EXCLURE de indemnites_break: (1) clauses de remboursement uniquement en cas de CESSION du bail ou du fonds — MEME si la clause mentionne un remboursement "au prorata temporis" ou des "mesures d'accompagnement", des lors que le fait generateur est une cession a un tiers (notamment hors du Groupe du Preneur) et NON l'exercice d'une option de conge/break par le preneur lui-meme. Exemple a EXCLURE: "remboursement des mesures d'accompagnement (franchise de loyer) accordees intuitu personae au prorata temporis en cas de cession du droit au bail ou du fonds de commerce au benefice d'une societe ne faisant pas partie du Groupe du Preneur" → ceci est une clause de cession, PAS une indemnite de break, NE JAMAIS lui attribuer de break_date ni l'inclure dans indemnites_break, meme partiellement; (2) penalites dues en cas de depart FAUTIF ou de resiliation anticipee HORS option de break (clause resolutoire, indemnite d'immobilisation); (3) indemnites dues entre deux dates de break. Inclure UNIQUEMENT les sommes dues lorsque le preneur EXERCE VALABLEMENT une option de break prevue au bail.
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

function computeBreaks(date_effet_str, date_fin_str, conditions_break_str, existing, duree_ferme_str) {
  const effet = parseFR(date_effet_str)
  const fin   = parseFR(date_fin_str)
  if (!effet || !fin) return existing || []

  const clauseText = (conditions_break_str || '').toLowerCase()

  const candidates = new Set()

  // Detect explicit waiver of ALL triennales (not partial "pour la durée ferme")
  // Partial waiver: "renonce...triennale...pour la durée ferme" → NOT a full waiver
  const hasWaiver = (() => {
    const basicWaiver = /renonce.{0,80}triennale|pas.{0,20}triennale|supprim.{0,20}triennale|faculté.{0,10}résiliation.{0,10}triennale/i.test(clauseText)
    if (!basicWaiver) return false
    // If the waiver is limited to the firm period, it's not a full waiver
    const isPartial = /pour la dur[eé]e ferme|pendant la dur[eé]e ferme|pour la p[eé]riode ferme/i.test(clauseText)
    return !isPartial
  })()

  // Parse duree_ferme into years+months
  const parseDureeFerme = (str) => {
    if (!str) return null
    const ymatch = String(str).match(/(\d+)\s*ans?/)
    const mmatch = String(str).match(/(\d+)\s*mois/)
    const years  = ymatch ? parseInt(ymatch[1]) : 0
    const months = mmatch ? parseInt(mmatch[1]) : 0
    return (years > 0 || months > 0) ? { years, months } : null
  }
  const dureeFerme = parseDureeFerme(duree_ferme_str)

  // Detect explicit single break dates
  const monthMap = { janvier:1,février:2,mars:3,avril:4,mai:5,juin:6,juillet:7,août:8,septembre:9,octobre:10,novembre:11,décembre:12 }

  // Detect "Nème période triennale pour la première fois" → first break at N×3 years
  const periodeTriennalePattern = /(deuxi[eè]me|troisi[eè]me|quatri[eè]me|2[eè]me|3[eè]me|4[eè]me|2e|3e|4e)\s+p[eé]riode\s+triennale/i
  const periodeMatch = periodeTriennalePattern.exec(clauseText)
  if (periodeMatch) {
    const ordinal = periodeMatch[1].toLowerCase()
    const n = ordinal.startsWith('deuxi') || ordinal.startsWith('2') ? 2
            : ordinal.startsWith('troisi') || ordinal.startsWith('3') ? 3
            : ordinal.startsWith('quatri') || ordinal.startsWith('4') ? 4 : null
    if (n) {
      // First break at N×3 years, puis toutes les 3 ans jusqu'à date_fin.
      // (Le "20" est un simple garde-fou anti-boucle-infinie, pas une limite
      // métier — avant, un plafond fixe à 12 ans coupait trop tôt les baux
      // plus longs, ex: manquant la break à 2035 sur un bail de 12+ ans.)
      for (let i = n; i < 20; i++) {
        const d = addYearsExpiry(effet, i * 3)
        if (d >= fin) break
        candidates.add(fmtFR(d))
      }
    }
  }

  if (!hasWaiver && !periodeMatch) {
    // Le droit de résiliation triennale (art. L.145-4 Code de commerce) est la
    // règle par défaut d'un bail commercial — on l'applique donc SYSTÉMATIQUEMENT
    // dès qu'aucune renonciation totale n'est détectée, plutôt que de dépendre
    // de la présence littérale du mot "triennale" dans le texte extrait (trop
    // fragile : une formulation différente, ou un extrait _sources incomplet,
    // faisait échouer silencieusement toute la détection).
    {
      // Une "duree ferme" incoherente (>= duree totale du bail, sans qu'un
      // renoncement total ait ete detecte) trahit typiquement une erreur
      // d'extraction — pas une vraie donnee bloquant tout droit de sortie.
      // On ne s'y fie que si elle laisse une marge reelle avant date_fin.
      const totalSpanYears = (fin - effet) / (1000 * 60 * 60 * 24 * 365.25)
      const dureeFermeFiable = dureeFerme && dureeFerme.years > 3 && dureeFerme.years < totalSpanYears - 0.5

      // If duree_ferme > 3 ans, first break starts at duree_ferme (not year 3)
      // This handles "renonce à la 1ère triennale, ferme 6 ans → break à 6 ans puis tous les 3 ans"
      if (dureeFermeFiable) {
        // First break at duree_ferme (peut ne pas être un multiple de 3, ex:
        // durée ferme 7 ans → premier congé possible "au plus tôt" à l'année 7).
        const firstBreak = new Date(effet.getFullYear() + dureeFerme.years, effet.getMonth() + (dureeFerme.months || 0), effet.getDate() - 1)
        if (firstBreak > effet && firstBreak < fin) candidates.add(fmtFR(firstBreak))
        // Le droit triennal reste ancré sur le rythme D'ORIGINE (années 3,6,9,12...
        // depuis date_effet), PAS un nouveau cycle qui repartirait de duree_ferme.
        // Ex: ferme 7 ans (renonciation aux échéances 3 et 6) → prochaine échéance
        // = année 9 (prochain multiple de 3 après 7), PAS année 7+3=10.
        let nextTri = Math.ceil(dureeFerme.years / 3) * 3
        if (nextTri <= dureeFerme.years) nextTri += 3 // si duree_ferme est déjà un multiple de 3, passer au suivant
        for (let y = nextTri; y < nextTri + 60; y += 3) { // garde-fou anti-boucle, pas une limite métier
          const d = addYearsExpiry(effet, y)
          if (d >= fin) break
          candidates.add(fmtFR(d))
        }
      } else {
        // Standard: tous les 3 ans à partir de l'année 3, jusqu'à date_fin
        // (aussi le repli si duree_ferme semble incohérente — cf. ci-dessus)
        for (let i = 1; i < 20; i++) {
          const y = i * 3
          const d = addYearsExpiry(effet, y)
          if (d >= fin) break
          candidates.add(fmtFR(d))
        }
      }
    }

    // Detect Nème année patterns — mais JAMAIS si ce "Nème année" est mentionné
    // dans un contexte de RENONCIATION ("renonçant... à l'expiration de la
    // 3ème et de la 6ème Année") : dans ce cas, c'est précisément l'échéance
    // qu'on retire, pas une option de sortie accordée. Sans ce garde-fou, ce
    // bloc générique ajoutait à tort les années explicitement renoncées.
    const yearPattern = /(\d+)[eè][mr]?[eè]?\s+ann[eé]e/gi
    let match
    while ((match = yearPattern.exec(clauseText)) !== null) {
      const n = parseInt(match[1])
      if (n > 0 && n < 12) {
        const contextBefore = clauseText.slice(Math.max(0, match.index - 100), match.index)
        const isRenounced = /renonc|sans facult|ne pourra|supprim/i.test(contextBefore)
        if (isRenounced) continue
        const d = addYearsExpiry(effet, n)
        if (d < fin) candidates.add(fmtFR(d))
      }
    }
  }

  // Always detect explicit dates in clause text
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

  // If nothing computed and no waiver, try duree_ferme as first break
  if (!candidates.size && !hasWaiver) {
    if (dureeFerme) {
      const breakDate = new Date(effet.getFullYear() + dureeFerme.years, effet.getMonth() + (dureeFerme.months || 0), effet.getDate() - 1)
      if (breakDate > effet && breakDate < fin) candidates.add(fmtFR(breakDate))
    }
    if (!candidates.size) return existing || []
  }

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

// Les break_options doivent être des dates pures "jj/mm/aaaa". Si le modèle
// a malgré tout renvoyé une phrase descriptive contenant une date, on extrait
// cette date plutôt que de perdre l'information silencieusement.
function sanitizeBreakDates(arr) {
  if (!Array.isArray(arr)) return arr
  return arr.map(b => {
    if (typeof b !== 'string') return null
    const s = b.trim()
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s
    const m = s.match(/(\d{1,2}\/\d{1,2}\/\d{4})/)
    return m ? m[1] : null
  }).filter(Boolean)
}

function sanitizeExtracted(data) {
  if (!data || typeof data !== 'object') return data
  const d = { ...data }
  d.break_options = sanitizeBreakDates(ensureArray(d.break_options))
  // Enrichir les breaks par calcul côté code — fiable à 100%, contrairement à
  // l'IA qui peut être incomplète (ex: s'arrêter après la 1ère échéance
  // triennale au lieu de continuer tous les 3 ans). On calcule TOUJOURS les
  // dates théoriques et on les FUSIONNE avec ce que l'IA a trouvé (union,
  // dédupliquée, triée) — plutôt que de ne s'en servir qu'en absence totale
  // de résultat de l'IA.
  if (d.date_effet && d.date_fin) {
    const computed = computeBreaks(d.date_effet, d.date_fin, d.conditions_break, [], d.duree_ferme)
    if (computed.length > 0) {
      const existing = new Set((d.break_options || []).map(b => b.trim()))
      const merged = [...(d.break_options || [])]
      computed.forEach(c => { if (!existing.has(c)) { merged.push(c); existing.add(c) } })
      merged.sort((a, b) => { const da = parseFR(a), db = parseFR(b); return (da && db) ? da - db : 0 })
      d.break_options = merged
    } else if (!d.break_options || d.break_options.length === 0) {
      d.break_options = []
    }
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
    d.champs_modifies.break_options      = sanitizeBreakDates(ensureArray(d.champs_modifies.break_options))
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
async function callClaude(base64, mediaType, prompt, timeoutMs = 120000) {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  try {
    const res = await fetch('https://vmtmwsbebzkwxfkdpqky.supabase.co/functions/v1/hyper-action', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      signal: controller.signal,
      body: JSON.stringify({
        model: 'claude-sonnet-4-6', max_tokens: 4096,
        messages: [{ role: 'user', content: [
          { type: 'document', source: { type: 'base64', media_type: mediaType, data: base64 } },
          { type: 'text', text: prompt }
        ]}]
      })
    })
    clearTimeout(timer)
    if (!res.ok) throw new Error('Erreur API: ' + res.status)
    const data = await res.json()
    if (data.type === 'error') {
      const msg = data.error?.message || 'Erreur Anthropic'
      if (msg.includes('100 PDF pages')) throw new Error('PDF > 100 pages : retirez les annexes avant de déposer.')
      if (msg.includes('too large') || msg.includes('file size') || msg.includes('32 MB')) throw new Error('Fichier trop volumineux (> 32 Mo) : compressez le PDF avant de déposer.')
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
  } catch (e) {
    clearTimeout(timer)
    if (e.name === 'AbortError') throw new Error('Timeout (> 2 min) : fichier trop lourd ou API indisponible.')
    throw e
  }
}

// Dual-pass extraction for bail: structural + financial in parallel, then merge

// ─── Sub-components ───────────────────────────────────────────────────────────

// Traverse a FileSystemEntry recursively, returning { file, path }
async function collectFiles(entry, path = '') {
  return new Promise(resolve => {
    if (entry.isFile) {
      entry.file(file => {
        const ext = file.name.split('.').pop().toLowerCase()
        if (['pdf', 'docx'].includes(ext)) {
          // Attach path info to file object
          const fileWithPath = Object.defineProperty(file, '_dirPath', { value: path, writable: true })
          resolve([fileWithPath])
        } else resolve([])
      }, () => resolve([]))
    } else if (entry.isDirectory) {
      const dirPath = path ? `${path}/${entry.name}` : entry.name
      const reader = entry.createReader()
      const results = []
      const readAll = () => {
        reader.readEntries(async entries => {
          if (!entries.length) {
            const nested = await Promise.all(results.map(e => collectFiles(e, dirPath)))
            resolve(nested.flat())
          } else {
            results.push(...entries)
            readAll()
          }
        }, () => resolve([]))
      }
      readAll()
    } else {
      resolve([])
    }
  })
}

// Group files by their directory path and infer bail/avenant links + actif groups
// Returns { autoLinks: { fileIndex -> bailFileIndex }, actifGroups: { fileIndex -> groupName } }
function inferLinksFromDirectories(allFiles) {
  const groups = {}
  allFiles.forEach((f, i) => {
    const dir = f._dirPath || ''
    if (!groups[dir]) groups[dir] = []
    groups[dir].push({ file: f, idx: i })
  })

  const autoLinks = {}   // avenant file index -> bail file index
  const actifGroups = {} // file index -> actif group name (top-level dir)

  // Determine top-level dir for each file (first path segment = bâtiment)
  allFiles.forEach((f, i) => {
    const parts = (f._dirPath || '').split('/').filter(Boolean)
    if (parts.length > 0) actifGroups[i] = parts[0].toUpperCase()
  })

  Object.values(groups).forEach(group => {
    if (group.length < 2) return
    const isAvenantName = f => /avenant|aven\b|addendum|protocole|rectif|modificat/i.test(f.file.name)
    const bails    = group.filter(g => !isAvenantName(g))
    const avenants = group.filter(g => isAvenantName(g))
    if (bails.length === 1 && avenants.length > 0) {
      avenants.forEach(av => { autoLinks[av.idx] = bails[0].idx })
    }
  })

  return { autoLinks, actifGroups }
}

function DropZone({ onFiles, disabled }) {
  const [dragging, setDragging] = useState(false)
  const [scanning, setScanning] = useState(false)
  const inputRef = useRef()

  const handle = useCallback(files => {
    const valid = Array.from(files).filter(f => ['pdf', 'docx'].includes(f.name.split('.').pop().toLowerCase()))
    if (valid.length) onFiles(valid)
    else alert('Format non supporté. PDF ou DOCX uniquement.')
  }, [onFiles])

  const handleDrop = useCallback(async e => {
    e.preventDefault()
    setDragging(false)
    const items = Array.from(e.dataTransfer.items || [])
    const hasEntries = items.some(i => i.webkitGetAsEntry)
    if (!hasEntries) { handle(e.dataTransfer.files); return }

    setScanning(true)
    const entries = items.map(i => i.webkitGetAsEntry()).filter(Boolean)
    const allFiles = (await Promise.all(entries.map(e => collectFiles(e, '')))).flat()
    const { autoLinks, actifGroups } = inferLinksFromDirectories(allFiles)
    setScanning(false)
    if (allFiles.length) onFiles(allFiles, autoLinks, actifGroups)
    else alert('Aucun fichier PDF ou DOCX trouvé dans le répertoire.')
  }, [handle, onFiles])

  return (
    <div
      className={`drop-zone${dragging ? ' dragging' : ''}${disabled ? ' disabled' : ''}`}
      onClick={() => !disabled && inputRef.current?.click()}
      onDragOver={e => { e.preventDefault(); setDragging(true) }}
      onDragLeave={e => { if (!e.currentTarget.contains(e.relatedTarget)) setDragging(false) }}
      onDrop={disabled ? undefined : handleDrop}
    >
      <input ref={inputRef} type="file" accept=".pdf,.docx" multiple style={{ display: 'none' }} onChange={e => handle(e.target.files)} />
      <div className="drop-icon">
        <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.2">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
          <polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>
        </svg>
      </div>
      <div className="drop-title">
        {scanning ? '⏳ Scan du répertoire en cours…' : 'Déposez fichiers ou répertoires ici'}
      </div>
      <div className="drop-sub">PDF ou DOCX · baux et avenants · répertoires et sous-répertoires acceptés</div>
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

function Field({ label, value, mono, verbose, full, source, item, pages, pageField }) {
  const safe = safeStr(value)
  const tooltip = source || (safe && safe !== 'Non renseigné' ? `Extrait : "${safe}"` : null)
  return (
    <div className={`field${verbose ? ' verbose-field' : ''}${full ? ' full' : ''}`}>
      <div className="field-lbl" style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
        {label}
        {tooltip && <span className="field-info" title={tooltip}>i</span>}
      </div>
      <div className={`field-val${!safe ? ' empty' : mono ? ' mono' : verbose ? ' verbose' : ''}`} style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
        {safe || 'Non renseigné'}
        {pageField && <PageJumpIcon item={item} pages={pages} field={pageField} />}
      </div>
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

function SurfaceTable({ surfaces, totalDeclared, totalLoyerDeclared }) {
  const safe = Array.isArray(surfaces) ? surfaces : []
  if (!safe.length) return null
  const isPark = r => { const cat = (r.categorie || r.typologie || '').toLowerCase(); return cat.includes('station') || cat.includes('parking') || cat.includes('place') }
  const mainRows = safe.filter(r => !isPark(r))
  const parkRows = safe.filter(r => isPark(r))
  const total = mainRows.reduce((acc, r) => acc + (parseFloat(String(r.surface_m2 || '').replace(/[^0-9.]/g, '')) || 0), 0)
  const totalLoyer = safe.reduce((acc, r) => acc + (parseAmount(r.loyer_annuel) || 0), 0)
  const parkTotalLoyer = parkRows.reduce((acc, r) => acc + (parseAmount(r.loyer_annuel) || 0), 0)
  const mainLoyerSum = mainRows.reduce((a, r) => a + (parseAmount(r.loyer_annuel) || 0), 0)
  // Si aucune ligne n'a de loyer propre (bail non ventilé par composante), on
  // retombe sur le loyer global du bail plutôt que d'afficher un total vide.
  const mainLoyerDisplay = mainLoyerSum > 0 ? mainLoyerSum : (parseAmount(totalLoyerDeclared) || 0)
  const mainLoyerIsFallback = mainLoyerSum === 0 && mainLoyerDisplay > 0
  // Écart entre la somme du détail et la surface totale déclarée du bail —
  // typiquement une quote-part de parties communes (« Surface Exploitée »,
  // SUBL, surface utile...) non ventilée ligne par ligne.
  const declared = parseFloat(String(totalDeclared || '').replace(',', '.')) || 0
  const commonAreaGap = declared > 0 && total > 0 ? declared - total : 0
  const hasNotableGap = Math.abs(commonAreaGap) > Math.max(1, declared * 0.01)

  // Compute unit price ONLY if both loyer_annuel AND surface_m2 are present
  const unitPrice = r => {
    if (r.prix_unitaire) return parseAmount(r.prix_unitaire)
    const loyer = parseAmount(r.loyer_annuel)
    const surf  = parseFloat(String(r.surface_m2 || '').replace(',', '.')) || 0
    // Only calculate if BOTH are explicitly provided
    if (loyer !== null && loyer > 0 && surf > 0) return Math.round(loyer / surf)
    return null // Cannot calculate - don't show
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
                    {mainLoyerDisplay > 0 ? fmtEur(mainLoyerDisplay) : '—'}
                    {mainLoyerIsFallback && <span title="Loyer global du bail — non ventilé par composante dans le document" style={{ fontSize: '10px', marginLeft: '3px', fontWeight: 400, fontStyle: 'italic', color: 'var(--text3)' }}>*</span>}
                  </td>
                </tr>
                {hasNotableGap && (
                  <tr>
                    <td colSpan={5} style={{ padding: '6px 10px 8px', fontSize: '11px', color: 'var(--text3)', fontStyle: 'italic', borderTop: 'none' }}>
                      {commonAreaGap > 0
                        ? <>Écart de <strong>+{commonAreaGap.toLocaleString('fr-FR', { maximumFractionDigits: 2 })} m²</strong> avec la surface totale louée ({declared.toLocaleString('fr-FR')} m²) — probablement une quote-part de parties communes non ventilée (bail parlant de « Surface Exploitée », SUBL, ou surface utile).</>
                        : <>La surface totale louée déclarée ({declared.toLocaleString('fr-FR')} m²) est inférieure de {Math.abs(commonAreaGap).toLocaleString('fr-FR', { maximumFractionDigits: 2 })} m² à la somme du détail ci-dessus — à vérifier.</>
                      }
                    </td>
                  </tr>
                )}
                {mainLoyerIsFallback && (
                  <tr>
                    <td colSpan={5} style={{ padding: '2px 10px 8px', fontSize: '10px', color: 'var(--text3)', fontStyle: 'italic', borderTop: 'none' }}>
                      * loyer global du bail, non ventilé par composante dans le document source
                    </td>
                  </tr>
                )}
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

// ─── État locatif : helpers de calcul ───────────────────────────────────────
function parseFrDate(s) {
  if (!s || typeof s !== 'string') return null
  const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/)
  if (!m) return null
  const d = new Date(+m[3], +m[2] - 1, +m[1])
  return isNaN(d.getTime()) ? null : d
}

function extractFloorInfo(adresse) {
  if (!adresse) return null
  const s = adresse.toLowerCase()
  if (/rez[\s-]*de[\s-]*chauss[ée]e|\brdc\b/.test(s)) return { key: 0, label: 'RDC' }
  let m = s.match(/(\d+)\s*(?:er|ère|e|ème|eme)\s*[ée]tage/)
  if (m) { const n = parseInt(m[1]); return { key: n, label: `${n}${n === 1 ? 'er' : 'e'} étage` } }
  m = s.match(/r\s*\+\s*(\d+)/)
  if (m) { const n = parseInt(m[1]); return { key: n, label: `R+${n}` } }
  return null
}

function monthsBetweenDates(a, b) { return (b - a) / (1000 * 60 * 60 * 24 * 30.44) }

function nextCriticalDate(t, today) {
  if (t.vacant) return null
  const candidates = [...(t.breaks || []), t.end].filter(d => d && d > today)
  if (!candidates.length) return null
  return candidates.sort((a, b) => a - b)[0]
}

function tenantStatus(t, today) {
  if (t.vacant) return 'vacant'
  const nc = nextCriticalDate(t, today)
  if (nc && monthsBetweenDates(today, nc) <= 18) return 'risk'
  return 'stable'
}

const SEGMENT_PALETTE = ['#14B8A6', '#6366F1', '#F59E0B', '#EC4899', '#0EA5E9', '#8B5CF6', '#F97316', '#10B981']
const ETAT_LOCATIF_COLORS = { stable: '#1D9E75', risk: '#EF9F27', vacant: '#B4B2A9' }
const ETAT_LOCATIF_BG     = { stable: '#EAF3DE', risk: '#FAEEDA', vacant: '#F1EFE8' }

// ─── Contrôle qualité : audit heuristique des baux déjà extraits ────────────
// Ne modifie rien — repère juste des cas suspects à vérifier/réextraire manuellement.
function parseYearsFromDureeText(s) {
  const m = String(s || '').match(/(\d+)\s*ans?/i)
  return m ? parseInt(m[1]) : null
}

// Une date de break tombant EXACTEMENT au jour anniversaire (sans le -1 jour
// conventionnel) suggère que l'IA n'a pas appliqué la convention standard des
// baux commerciaux — sauf si le bail écrit cette date explicitement en toutes
// lettres, auquel cas c'est normal et cette détection est un faux positif à
// écarter au cas par cas.
function matchesExactAnniversary(breakDateStr, effet) {
  const bd = parseFrDate(breakDateStr)
  if (!bd || !effet) return null
  for (let n = 1; n < 40; n++) {
    const exact = new Date(effet.getFullYear() + n, effet.getMonth(), effet.getDate())
    if (bd.getTime() === exact.getTime()) return n
  }
  return null
}

function auditBail(row) {
  const d = row.data || {}
  const issues = []
  const building = d.immeuble || d.adresse || row.file_name
  const tenant = shortPartyName(d.preneur)
  const label = tenant ? `${building} — ${tenant}` : building

  const effet = parseFrDate(d.date_effet)
  const fin = parseFrDate(d.date_fin)
  const dureeFermeYears = parseYearsFromDureeText(d.duree_ferme)
  const dureeTotaleYears = parseYearsFromDureeText(d.duree_totale)
  const rawBreaks = Array.isArray(d.break_options) ? d.break_options : []
  const cleanBreaks = rawBreaks.filter(b => typeof b === 'string' && /^\d{2}\/\d{2}\/\d{4}$/.test(b.trim()))

  // 1. Breaks triennaux potentiellement incomplets (le bug qu'on vient de corriger)
  if (effet && fin && dureeFermeYears && dureeFermeYears < (dureeTotaleYears || 99)) {
    const expected = []
    let y = dureeFermeYears
    while (true) {
      const dt = new Date(effet.getFullYear() + y, effet.getMonth(), effet.getDate() - 1)
      if (dt >= fin) break
      expected.push(dt)
      y += 3
    }
    if (expected.length > cleanBreaks.length) {
      issues.push({
        type: 'breaks_incomplets',
        severity: 'high',
        detail: `${cleanBreaks.length} break(s) enregistré(s), ${expected.length} attendu(s) si la périodicité triennale continue jusqu'à l'échéance (${expected.map(fmtFR).join(', ')})`,
      })
    }
  }

  // 1bis. Résidu de l'ancien bug de calcul (durée ferme non multiple de 3) :
  // avant correction, la 2e échéance était calculée à durée_ferme+3 ans au lieu
  // du prochain multiple de 3 depuis la date d'effet — repère les bails ayant
  // potentiellement cette date incorrecte figée en base suite à une extraction
  // antérieure au correctif.
  if (effet && fin && dureeFermeYears && dureeFermeYears % 3 !== 0) {
    const wrongYear = dureeFermeYears + 3
    const correctYear = Math.ceil(dureeFermeYears / 3) * 3
    const wrongStr = fmtFR(new Date(effet.getFullYear() + wrongYear, effet.getMonth(), effet.getDate() - 1))
    if (cleanBreaks.includes(wrongStr)) {
      const correctStr = fmtFR(new Date(effet.getFullYear() + correctYear, effet.getMonth(), effet.getDate() - 1))
      issues.push({
        type: 'break_ancien_calcul',
        severity: 'high',
        detail: `Break enregistrée au ${wrongStr} (durée ferme + 3 ans) — probablement un résidu de l'ancien calcul, avant correction. La bonne échéance devrait être le ${correctStr} (prochain multiple de 3 depuis la date d'effet). Réextraction recommandée pour purger la date incorrecte.`,
      })
    }
  }

  // 2. break_options mal formés (texte descriptif au lieu d'une date pure)
  const malformed = rawBreaks.filter(b => typeof b === 'string' && !/^\d{2}\/\d{2}\/\d{4}$/.test(b.trim()))
  if (malformed.length > 0) {
    issues.push({
      type: 'break_format',
      severity: 'medium',
      detail: `${malformed.length} entrée(s) de break_options mal formatée(s) (texte au lieu d'une date pure) — récupérée(s) à l'affichage si une date y est repérable, mais à vérifier`,
    })
  }

  // 2ter. Champs date/durée pollués par du texte concaténé (ex: date_fin =
  // "31/12/2023 (renouvelable, terme absolu 31/12/2034)" au lieu d'une date
  // pure) — résidu d'extractions antérieures à l'ajout du champ dédié
  // reconduction_tacite.date_limite_absolue. La date extraite reste
  // exploitable (parseFrDate prend la première trouvée) mais l'info de
  // plafond est perdue pour le Gantt tant que le bail n'est pas réextrait.
  ;[['date_fin', d.date_fin], ['date_effet', d.date_effet], ['date_signature', d.date_signature]].forEach(([field, val]) => {
    if (typeof val === 'string' && val.trim() && !/^\d{2}\/\d{2}\/\d{4}$/.test(val.trim())) {
      issues.push({
        type: 'champ_date_pollue',
        severity: 'high',
        detail: `${field} = "${val}" contient du texte en plus de la date — devrait être une date pure (JJ/MM/AAAA). Réextraction recommandée pour isoler l'information annexe (ex: plafond de reconduction) dans son propre champ.`,
      })
    }
  })
  if (typeof d.duree_totale === 'string' && d.duree_totale.trim() && !/^\d+\s*ans?$/i.test(d.duree_totale.trim())) {
    issues.push({
      type: 'champ_duree_pollue',
      severity: 'high',
      detail: `duree_totale = "${d.duree_totale}" contient du texte en plus de la durée — devrait être une valeur simple (ex: "1 an", "9 ans"). Réextraction recommandée.`,
    })
  }

  // 2bis. Break(s) tombant exactement au jour anniversaire de date_effet, sans
  // le "-1 jour" conventionnel des baux commerciaux français — sauf si le bail
  // écrit cette date en toutes lettres (auquel cas c'est un faux positif).
  if (effet) {
    const anniversaryHits = cleanBreaks
      .map(b => ({ b, n: matchesExactAnniversary(b, effet) }))
      .filter(x => x.n !== null)
    if (anniversaryHits.length > 0) {
      issues.push({
        type: 'break_jour_anniversaire',
        severity: 'medium',
        detail: `${anniversaryHits.map(x => `${x.b} (année ${x.n})`).join(', ')} tombe(nt) exactement au jour anniversaire de la prise d'effet — devrait être le jour précédent (convention "jour anniversaire moins 1 jour"), sauf si cette date est écrite explicitement en toutes lettres dans le bail. À vérifier sur le document source.`,
        affectedDates: anniversaryHits.map(x => x.b),
      })
    }
  }

  // 3. Incohérence entre surface_totale_m2 (utilisé par le Dashboard) et la
  // somme des lignes de surfaces_detail (utilisée par l'État locatif) — un
  // écart signale soit une ligne manquante/mal classée, soit un cas légitime
  // "Surface Exploitée" (surface totale incluant une quote-part de parties
  // communes, distincte de la somme des composantes individuellement louées)
  // — on ne peut pas distinguer les deux depuis les données déjà extraites,
  // donc on signale à vérifier sans présumer d'une erreur.
  const detailSum = (Array.isArray(d.surfaces_detail) ? d.surfaces_detail : [])
    .filter(r => !(r.categorie || '').toLowerCase().includes('station'))
    .reduce((a, r) => a + (parseFloat(String(r.surface_m2 || '').replace(',', '.')) || 0), 0)
  const totalM2 = parseFloat(String(d.surface_totale_m2 || '').replace(',', '.')) || 0
  if (detailSum > 0 && totalM2 > 0) {
    const gap = Math.abs(detailSum - totalM2)
    if (gap > Math.max(2, totalM2 * 0.02)) {
      issues.push({
        type: 'surface_incoherente',
        severity: 'medium',
        detail: `Surface totale déclarée ${totalM2} m², somme des lignes de surfaces_detail = ${detailSum.toFixed(2)} m² (écart de ${gap.toFixed(2)} m²) — soit une ligne manquante/mal classée, soit un cas légitime "Surface Exploitée" (incluant une quote-part de parties communes) : à vérifier sur le document source`,
      })
    }
  }

  // 4. Durée ferme quasi égale à la durée totale — ambigu par nature : peut
  // être une vraie renonciation totale à toute sortie anticipée (légitime),
  // ou une erreur d'extraction (duree_ferme mal déduite faute de restriction
  // explicite dans le bail). Impossible de trancher depuis les données seules
  // — signalé pour vérification, pas comme une erreur certaine.
  if (dureeFermeYears && dureeTotaleYears && dureeFermeYears >= dureeTotaleYears - 1) {
    issues.push({
      type: 'duree_ferme_ambigue',
      severity: 'low',
      detail: `Durée ferme (${dureeFermeYears} ans) quasi égale à la durée totale (${dureeTotaleYears} ans) — soit le bail prévoit vraiment une renonciation totale à toute sortie anticipée, soit c'est une erreur d'extraction (rien n'est explicitement restreint dans le texte). À vérifier sur le document source.`,
    })
  }

  // Note : la reconduction tacite (art. L145-9 Code de commerce) est le régime
  // légal par défaut pour la quasi-totalité des baux commerciaux — son absence
  // du texte n'est pas anormale et ne justifie pas une vérification systématique.

  return { row, label, issues, dismissed: !!d._qc_dismissed }
}

// Détection de doublons probables : deux baux (ou plus) portant sur le même
// actif ET le même preneur ET une surface quasi identique sont très probablement
// le même document déposé/extrait deux fois. Nécessite de comparer TOUS les
// baux entre eux — contrairement aux autres contrôles, qui n'examinent qu'un
// bail à la fois.
function findDuplicateBails(bails) {
  const groups = {}
  bails.forEach(row => {
    const d = row.data || {}
    const building = (d.immeuble || d.adresse || '').toLowerCase().trim()
    const tenant = (shortPartyName(d.preneur) || '').toLowerCase().trim()
    if (!building || !tenant) return
    const key = `${building}|||${tenant}`
    if (!groups[key]) groups[key] = []
    groups[key].push(row)
  })

  const detailsByRowId = {}
  Object.values(groups).forEach(group => {
    if (group.length < 2) return
    // Sous-regroupement par surface (arrondie au m² près) pour éviter les faux
    // positifs entre deux baux légitimement différents sur le même actif/preneur
    // (ex: extension par avenant réextraite comme bail séparé, surfaces différentes).
    const bySurface = {}
    group.forEach(row => {
      const s = parseFloat(String(row.data?.surface_totale_m2 || '').replace(',', '.'))
      const sKey = !isNaN(s) && s > 0 ? Math.round(s) : `sans-surface-${row.id}` // pas de surface → jamais groupé avec un autre
      if (!bySurface[sKey]) bySurface[sKey] = []
      bySurface[sKey].push(row)
    })
    Object.entries(bySurface).forEach(([sKey, rows]) => {
      if (rows.length < 2) return
      rows.forEach(row => {
        const others = rows.filter(r => r.id !== row.id).map(r => r.file_name).join(', ')
        detailsByRowId[row.id] = `Semble être un doublon — même actif, même preneur, surface quasi identique (~${sKey} m²) que : ${others}`
      })
    })
  })
  return detailsByRowId
}

// Fusionne les modifications des avenants (triés chronologiquement) sur les
// données du bail — sans ça, un bail dont la date d'effet n'était pas fixée à
// la signature (condition suspensive, VEFA...) et fixée ensuite par avenant
// afficherait des dates vides ou erronées dans l'état locatif.
function mergedBailData(row) {
  const base = { ...(row.data || {}) }
  const toSortable = s => {
    const m = String(s || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/)
    return m ? `${m[3]}-${m[2]}-${m[1]}` : String(s || '')
  }
  const avs = [...(row.avenants || [])].sort((a, b) => {
    const ka = toSortable(a.data?.date_effet_avenant || a.data?.date_signature_avenant || a.created_at)
    const kb = toSortable(b.data?.date_effet_avenant || b.data?.date_signature_avenant || b.created_at)
    return ka.localeCompare(kb)
  })
  avs.forEach(av => {
    const mods = av.data?.champs_modifies || {}
    Object.entries(mods).forEach(([k, v]) => {
      if (v !== null && v !== undefined && !(Array.isArray(v) && v.length === 0) && k !== '_sources') {
        base[k] = v
        // Un avenant qui redéfinit explicitement break_options rend cette
        // valeur définitive — le calcul de secours (computeBreaks), basé sur
        // la clause du bail INITIAL, ne doit plus venir y rajouter les
        // anciennes échéances superseded par cet avenant.
        if (k === 'break_options') base._breakOptionsFromAvenant = true
      }
    })
    // surfaces_apres (champ séparé, hors champs_modifies) reflète l'assiette
    // EXACTE post-avenant — prioritaire sur champs_modifies.surfaces_detail,
    // qui lui n'est presque jamais renseigné pour une simple extension.
    if (Array.isArray(av.data?.surfaces_apres) && av.data.surfaces_apres.length > 0) {
      base.surfaces_detail = av.data.surfaces_apres
    }
  })
  return base
}

function EtatLocatifModal({ building, bails, onClose }) {
  const [tooltip, setTooltip] = useState(null) // { x, y, tenant }
  const today = new Date()

  // Une ligne par BAIL (pas par étage) — un bail qui occupe plusieurs étages
  // affiche un libellé de localisation agrégé plutôt que d'être éclaté sur
  // plusieurs lignes. Un même preneur avec plusieurs baux distincts garde
  // autant de lignes que de baux (jamais fusionné).
  const bailRows = useMemo(() => {
    const rows = bails.map(row => {
      const d = mergedBailData(row)

      // Si un avenant confirme la date d'effet (levée de condition suspensive),
      // on recalcule date_fin et break_options à partir d'elle AVANT la
      // détection "estimée" ci-dessous — une fois la date de départ certaine,
      // ce n'est plus une estimation incertaine (contrairement au cas VEFA
      // sans confirmation), donc pas de pointillés à afficher pour ce bail.
      if (Array.isArray(row.avenants) && row.avenants.length > 0) {
        const toSortableAv = s => { const m = String(s || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/); return m ? `${m[3]}-${m[2]}-${m[1]}` : String(s || '') }
        const sortedAvs = [...row.avenants].sort((a, b) =>
          toSortableAv(a.data?.date_effet_avenant || a.data?.date_signature_avenant || a.created_at)
            .localeCompare(toSortableAv(b.data?.date_effet_avenant || b.data?.date_signature_avenant || b.created_at)))
        let confirmedByAvenant = false
        let hasExplicitDateFin = false
        let hasExplicitBreaks = false
        sortedAvs.forEach(av => {
          const mods = av.data?.champs_modifies || {}
          if (mods.date_effet) confirmedByAvenant = true
          if (mods.date_fin) hasExplicitDateFin = true
          if (Array.isArray(mods.break_options) && mods.break_options.length > 0) hasExplicitBreaks = true
        })
        if (confirmedByAvenant) {
          const startConfirmed = parseFrDate(d.date_effet)
          if (startConfirmed) {
            if (!hasExplicitDateFin) {
              const m = String(d.duree_totale || '').match(/(\d+)\s*ans?/i)
              if (m) {
                const end = new Date(startConfirmed.getFullYear() + parseInt(m[1]), startConfirmed.getMonth(), startConfirmed.getDate() - 1)
                d.date_fin = fmtFR(end)
              }
            }
            if (!hasExplicitBreaks) {
              const recalc = computeBreaks(d.date_effet, d.date_fin, d.conditions_break, [], d.duree_ferme)
              if (recalc.length > 0) d.break_options = recalc
            }
          }
        }
      }

      let start = parseFrDate(d.date_effet)
      let end = parseFrDate(d.date_fin)
      let estimated = false
      let estimatedField = null // 'end' | 'start' — précise quelle date a été calculée
      // Cas VEFA : la date de fin est souvent formulée en relatif
      // ("9 ans à compter de la Date de Livraison") plutôt qu'en date fixe.
      // Si on a une date de départ (même prévisionnelle) et une durée totale,
      // on calcule une échéance estimée plutôt que de renoncer à afficher le bail.
      if (start && !end) {
        const m = String(d.duree_totale || '').match(/(\d+)\s*ans?/i)
        if (m) {
          end = new Date(start)
          end.setFullYear(end.getFullYear() + parseInt(m[1]))
          estimated = true
          estimatedField = 'end'
        }
      } else if (end && !start) {
        // Cas symétrique : date de fin connue mais date d'effet absente de
        // l'extraction (champ manquant, pas forcément un VEFA) — on la
        // recalcule en remontant depuis la date de fin et la durée totale,
        // plutôt que de perdre le bail dans le filet "dates non déterminées".
        const m = String(d.duree_totale || '').match(/(\d+)\s*ans?/i)
        if (m) {
          start = new Date(end.getFullYear() - parseInt(m[1]), end.getMonth(), end.getDate() + 1)
          estimated = true
          estimatedField = 'start'
        }
      }
      // Fusionne (union) les breaks déjà stockés avec le calcul déterministe —
      // rattrape à l'affichage les extractions antérieures au renforcement de
      // computeBreaks (ex: 3e échéance triennale manquante en base). SAUF si un
      // avenant a explicitement redéfini break_options : dans ce cas cette
      // valeur est définitive, le calcul basé sur la clause du bail INITIAL ne
      // doit pas venir y rajouter les anciennes échéances superseded.
      const storedBreaks = sanitizeBreakDates(d.break_options || [])
      const computedBreaksEL = d._breakOptionsFromAvenant ? [] : computeBreaks(d.date_effet, d.date_fin, d.conditions_break, [], d.duree_ferme)
      const mergedBreaksSet = new Set(storedBreaks.map(b => b.trim()))
      computedBreaksEL.forEach(c => mergedBreaksSet.add(c))
      const mergedBreaks = [...mergedBreaksSet].sort((a, b) => { const da = parseFR(a), db = parseFR(b); return (da && db) ? da - db : 0 })

      // Localisation : liste des niveaux distincts occupés par ce bail (issus
      // de surfaces_detail), triés du plus bas au plus haut, sinon repli sur
      // l'adresse/l'immeuble.
      const detailRows = (d.surfaces_detail || []).filter(r =>
        (r.niveau || r.localisation) && !(r.categorie || '').toLowerCase().includes('station')
      )
      let locationLabel, sortKey, surface
      if (detailRows.length > 0) {
        const niveaux = [...new Set(detailRows.map(r => r.niveau || r.localisation))]
        // On n'affiche que l'étage reconnu (RDC, 1er, 2e...), pas le texte brut
        // du bail qui peut contenir des mentions de lot très variables en
        // longueur ("galerie lot 1 voir plan...") — bien plus lisible en badge
        // compact, et le texte complet reste visible dans l'infobulle.
        const infoList = niveaux
          .map(n => ({ raw: n, info: extractFloorInfo(n) }))
          .sort((a, b) => (a.info?.key ?? 9999) - (b.info?.key ?? 9999))
        locationLabel = infoList.map(x => x.info ? x.info.label : x.raw).join(', ')
        sortKey = infoList[0]?.info?.key ?? 9999
        // Un seul étage occupé : la "surface totale" du bail (surface exploitée
        // le cas échéant, incluant la quote-part de parties communes) s'applique
        // entièrement à cet étage — on la préfère à la ligne de détail, qui peut
        // ne représenter que le lot de bureaux hors quote-part.
        const totalM2 = parseFloat(String(d.surface_totale_m2 || '').replace(',', '.')) || 0
        surface = (detailRows.length === 1 && totalM2 > 0)
          ? totalM2
          : detailRows.reduce((a, r) => a + (parseFloat(String(r.surface_m2 || '').replace(',', '.').replace(/[^\d.]/g, '')) || 0), 0)
      } else {
        const info = extractFloorInfo(d.adresse)
        locationLabel = info ? info.label : (d.immeuble || d.adresse || 'Lot non localisé')
        sortKey = info ? info.key : 9999
        surface = parseFloat(String(d.surface_totale_m2 || '').replace(',', '.')) || 0
      }

      return {
        sortKey,
        tenant: {
          name: shortPartyName(d.preneur) || row.file_name,
          start, end, estimated, estimatedField,
          breaks: mergedBreaks.map(parseFrDate).filter(Boolean),
          loyer: parseAmount ? parseAmount(d.loyer_signature_montant) : (parseFloat(String(d.loyer_signature_montant || '').replace(/[^\d.,]/g, '').replace(',', '.')) || null),
          reconductionTacite: d.reconduction_tacite?.applicable ? {
            preavis: d.reconduction_tacite.preavis || null,
            periodicite: d.reconduction_tacite.periodicite || null,
            dateLimiteAbsolue: parseFrDate(d.reconduction_tacite.date_limite_absolue) || null,
          } : null,
          row,
          locationLabel,
          surface,
          showRent: true, // un bail = une ligne désormais, plus de risque de doublon d'affichage du loyer
        },
      }
    })
    const resolved = rows.filter(r => r.sortKey !== 9999).sort((a, b) => b.sortKey - a.sortKey)
    const unresolved = rows.filter(r => r.sortKey === 9999)
    return [...resolved, ...unresolved].map(r => r.tenant)
  }, [bails])

  const allTenants = bailRows
  const withDates = allTenants.filter(t => t.start && t.end)
  // Segment "reconduction tacite" : la barre se prolonge jusqu'à la première
  // date de sortie effective possible à partir d'aujourd'hui — c-a-d aujourd'hui
  // + préavis (souvent 6 mois), puis effet reporté au dernier jour du trimestre
  // civil en cours à cette date (règle usuelle du congé en tacite reconduction).
  function reconductionCutoff(t) {
    if (!t.reconductionTacite) return t.end
    // Plafond contractuel absolu (ex: "ne pourra excéder 12 ans, fin
    // automatique le 31/12/2034 sans formalité") : la reconduction tacite ne
    // peut jamais dépasser cette date, quel que soit le préavis calculé
    // depuis aujourd'hui — ce plafond prime sur le calcul par défaut.
    if (t.reconductionTacite.dateLimiteAbsolue) {
      return t.reconductionTacite.dateLimiteAbsolue > t.end ? t.reconductionTacite.dateLimiteAbsolue : t.end
    }
    const m = String(t.reconductionTacite.preavis || '').match(/(\d+)\s*mois/i)
    const noticeMonths = m ? parseInt(m[1]) : 6
    const target = new Date(today)
    target.setMonth(target.getMonth() + noticeMonths)
    const quarterEndMonth = Math.floor(target.getMonth() / 3) * 3 + 2 // 2=mars,5=juin,8=sept,11=déc
    const cutoff = new Date(target.getFullYear(), quarterEndMonth + 1, 0) // dernier jour du trimestre
    return cutoff > t.end ? cutoff : t.end // jamais avant le terme ferme lui-même
  }
  const effectiveEnd = t => reconductionCutoff(t)
  const domainStart = withDates.length ? new Date(Math.min(...withDates.map(t => t.start)) - 1000 * 60 * 60 * 24 * 180) : new Date(today.getFullYear() - 1, 0, 1)
  const domainEnd = withDates.length ? new Date(Math.max(...withDates.map(t => effectiveEnd(t))) + 1000 * 60 * 60 * 24 * 180) : new Date(today.getFullYear() + 5, 0, 1)
  const domainMs = domainEnd - domainStart
  const allYears = []
  for (let y = domainStart.getFullYear(); y <= domainEnd.getFullYear(); y++) allYears.push(y)
  // Sur une étendue large (beaucoup d'années), répartir un nombre limité
  // d'étiquettes UNIFORMÉMENT sur toute la largeur (premier et dernier inclus)
  // plutôt qu'un espacement fixe + ajout forcé du dernier — qui pouvait coller
  // les deux dernières années l'une contre l'autre quand le pas ne tombait pas
  // juste en fin de plage.
  const maxLabels = allYears.length > 24 ? 8 : allYears.length > 13 ? 11 : allYears.length
  const years = allYears.length <= maxLabels || maxLabels <= 1
    ? allYears
    : [...new Set(Array.from({ length: maxLabels }, (_, i) =>
        allYears[Math.round(i * (allYears.length - 1) / (maxLabels - 1))]))]

  function fmt(d) { return d ? d.toLocaleDateString('fr-FR') : '—' }

  // Découpe un bail en segments entre chaque break (= périodes contractuelles
  // successives), chacun coloré différemment via une palette cyclique.
  function tenantSegments(t) {
    if (!t.start || !t.end) return []
    const points = [t.start, ...t.breaks.filter(b => b > t.start && b < t.end).sort((a, b) => a - b), t.end]
    return points.slice(0, -1).map((p, i) => ({ start: p, end: points[i + 1], color: SEGMENT_PALETTE[i % SEGMENT_PALETTE.length] }))
  }

  function TenantTooltip({ x, y, tenant: t }) {
    const status = tenantStatus(t, today)
    const nc = nextCriticalDate(t, today)
    return (
      <div style={{
        position: 'fixed', left: x + 16, top: y - 12, zIndex: 3000, pointerEvents: 'none',
        background: 'var(--surface)', border: '1px solid var(--border2)', borderRadius: '10px',
        boxShadow: '0 10px 28px rgba(0,0,0,.22)', padding: '12px 14px', minWidth: '220px', maxWidth: '280px',
      }}>
        <div style={{ fontSize: '14px', fontWeight: 700, color: 'var(--text)', marginBottom: '8px' }}>{t.name}</div>
        {!t.start || !t.end ? (
          <div style={{ fontSize: '12px', color: 'var(--text3)', fontStyle: 'italic' }}>Dates non renseignées</div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '5px' }}>
            {t.surface > 0 && (
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px' }}>
                <span style={{ color: 'var(--text3)' }}>Surface</span>
                <span style={{ fontWeight: 600, color: 'var(--text)' }}>{Math.round(t.surface)} m²</span>
              </div>
            )}
            {t.loyer > 0 && (
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px' }}>
                <span style={{ color: 'var(--text3)' }}>Loyer HT/HC</span>
                <span style={{ fontWeight: 600, color: 'var(--text)' }}>{fmtEur(t.loyer)}</span>
              </div>
            )}
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px' }}>
              <span style={{ color: 'var(--text3)' }}>Prise d'effet{t.estimatedField === 'start' ? ' (estimée)' : ''}</span>
              <span style={{ fontWeight: 600, color: 'var(--text)' }}>{fmt(t.start)}</span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px' }}>
              <span style={{ color: 'var(--text3)' }}>{t.reconductionTacite ? 'Fin du terme ferme' : 'Échéance'}{t.estimatedField === 'end' ? ' (estimée)' : ''}</span>
              <span style={{ fontWeight: 600, color: 'var(--text)' }}>{fmt(t.end)}</span>
            </div>
            {t.estimated && (
              <div style={{ fontSize: '10px', color: 'var(--text3)', fontStyle: 'italic' }}>
                {t.estimatedField === 'start'
                  ? 'Date d\'effet non extraite — recalculée à partir de la date de fin et de la durée totale'
                  : 'Calculée à partir de la durée totale et de la date de livraison prévisionnelle (VEFA)'}
              </div>
            )}
            {t.reconductionTacite && (
              <div style={{ fontSize: '10px', color: 'var(--accent)', fontStyle: 'italic' }}>
                ↻ Reconduction tacite{t.reconductionTacite.periodicite ? ` ${t.reconductionTacite.periodicite}` : ''} au-delà de cette date
                {t.reconductionTacite.preavis ? ` (préavis ${t.reconductionTacite.preavis})` : ''}
                {t.reconductionTacite.dateLimiteAbsolue ? (
                  <><br />⛔ Fin automatique (plafond contractuel, sans formalité) : {fmt(t.reconductionTacite.dateLimiteAbsolue)}</>
                ) : (
                  <><br />Prochaine sortie possible : {fmt(reconductionCutoff(t))}</>
                )}
              </div>
            )}
            {t.breaks.length > 0 && (
              <div style={{ paddingTop: '5px', marginTop: '2px', borderTop: '1px solid var(--border)' }}>
                <div style={{ fontSize: '11px', color: 'var(--text3)', marginBottom: '4px' }}>Options de sortie</div>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '5px' }}>
                  {t.breaks.map((b, i) => (
                    <span key={i} style={{ fontSize: '11px', fontWeight: 600, padding: '2px 7px', borderRadius: '999px', background: SEGMENT_PALETTE[(i + 1) % SEGMENT_PALETTE.length] + '22', color: SEGMENT_PALETTE[(i + 1) % SEGMENT_PALETTE.length] }}>{fmt(b)}</span>
                  ))}
                </div>
              </div>
            )}
            {nc && (
              <div style={{
                marginTop: '4px', padding: '6px 8px', borderRadius: '6px', fontSize: '11px', fontWeight: 600,
                background: status === 'risk' ? ETAT_LOCATIF_BG.risk : 'var(--surface2)', color: status === 'risk' ? ETAT_LOCATIF_COLORS.risk : 'var(--text3)',
              }}>
                {status === 'risk' ? '⚠ ' : ''}Prochaine échéance dans {Math.round(monthsBetweenDates(today, nc))} mois
              </div>
            )}
          </div>
        )}
      </div>
    )
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" style={{ width: '95vw', height: '95vh', maxWidth: 'none', maxHeight: '95vh' }} onClick={e => e.stopPropagation()}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div className="modal-title">État locatif — {building}</div>
          <button onClick={onClose} title="Fermer" style={{ background: 'none', border: 'none', fontSize: '20px', lineHeight: 1, cursor: 'pointer', color: 'var(--text2)', padding: '4px' }}>✕</button>
        </div>

        <div style={{ overflowY: 'auto', flex: 1, paddingRight: '4px' }}>
          {bailRows.length === 0 ? (
            <div style={{ padding: '40px', textAlign: 'center', color: 'var(--text3)', fontSize: '13px' }}>
              Aucun bail rattaché à cet actif groupant pour le moment.
            </div>
          ) : (() => {
            const COL_W = [190, 75, 100, 100]
            return (
            <div style={{ border: '1px solid var(--border)', borderRadius: '12px', overflow: 'hidden' }}>
              <div style={{ display: 'flex', alignItems: 'stretch', height: '34px', borderBottom: '1px solid var(--border)', background: 'var(--surface2)' }}>
                {['Preneur', 'Surface', 'Loyer', 'Niveau'].map((h, i) => (
                  <div key={h} style={{ width: `${COL_W[i]}px`, flexShrink: 0, display: 'flex', alignItems: 'center', paddingLeft: '10px', fontSize: '10px', fontWeight: 700, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.04em' }}>{h}</div>
                ))}
                <div style={{ flex: 1, minWidth: 0, position: 'relative', overflow: 'hidden' }}>
                  {years.map(y => {
                    const yd = new Date(y, 0, 1)
                    const pct = ((yd - domainStart) / domainMs) * 100
                    // Éviter que l'étiquette de bord (première/dernière année)
                    // ne dépasse du cadre et se fasse couper/chevaucher — on la
                    // cale contre son trait plutôt que de la centrer dessus.
                    const align = pct < 4 ? 'left' : pct > 96 ? 'right' : 'center'
                    const transform = align === 'left' ? 'translateX(0)' : align === 'right' ? 'translateX(-100%)' : 'translateX(-50%)'
                    return (
                      <div key={y} style={{ position: 'absolute', left: `${pct}%`, top: 0, bottom: 0, borderLeft: '1px solid var(--border)' }}>
                        <span style={{ position: 'absolute', top: '9px', left: 0, transform, fontSize: '13px', fontWeight: 700, color: 'var(--text)' }}>{y}</span>
                      </div>
                    )
                  })}
                </div>
              </div>
              {bailRows.map(t => {
                const ROW_H = 52
                const status = t.start && t.end ? tenantStatus(t, today) : null
                return (
                  <div key={t.row.id} style={{ display: 'flex', alignItems: 'center', borderBottom: '1px solid var(--border)', height: `${ROW_H}px` }}>
                    <div
                      onClick={() => t.row && window.dispatchEvent(new CustomEvent('etatlocatif-select', { detail: t.row }))}
                      style={{ width: `${COL_W[0]}px`, flexShrink: 0, paddingLeft: '10px', paddingRight: '6px', fontSize: '13px', fontWeight: 600, color: 'var(--text)', cursor: t.row ? 'pointer' : 'default', display: 'flex', alignItems: 'center', gap: '5px', flexWrap: 'wrap', lineHeight: 1.3 }}>
                      <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{t.name}</span>
                      {status === 'risk' && <span title="Échéance de sortie dans moins de 18 mois" style={{
                        flexShrink: 0, fontSize: '9.5px', fontWeight: 700, padding: '1px 7px', borderRadius: '999px',
                        background: '#FAEEDA', color: '#B8860B', whiteSpace: 'nowrap',
                      }}>Échéance proche</span>}
                      {t.estimated && <span title={t.estimatedField === 'start' ? 'Date d\'effet non extraite — recalculée à partir de la date de fin et de la durée totale' : 'Échéance estimée (VEFA)'} style={{ flexShrink: 0, color: 'var(--text3)' }}>≈</span>}
                      {t.reconductionTacite && <span title={`Reconduction tacite${t.reconductionTacite.periodicite ? ' ' + t.reconductionTacite.periodicite : ''}${t.reconductionTacite.preavis ? ', préavis ' + t.reconductionTacite.preavis : ''}`} style={{ flexShrink: 0, color: 'var(--accent)' }}>↻</span>}
                    </div>
                    <div style={{ width: `${COL_W[1]}px`, flexShrink: 0, fontSize: '12px', color: 'var(--text2)', paddingLeft: '10px' }}>{t.surface > 0 ? `${Math.round(t.surface)} m²` : '—'}</div>
                    <div style={{ width: `${COL_W[2]}px`, flexShrink: 0, fontSize: '12px', color: 'var(--text2)', paddingLeft: '10px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{t.loyer > 0 ? fmtEur(t.loyer) : '—'}</div>
                    <div style={{ width: `${COL_W[3]}px`, flexShrink: 0, paddingLeft: '10px' }}>
                      <span title={t.locationLabel} style={{
                        fontSize: '11px', fontWeight: 600, color: 'var(--accent)', background: 'var(--accent-bg)',
                        padding: '3px 9px', borderRadius: '999px', display: 'inline-block', maxWidth: '100%',
                        overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                      }}>
                        {t.locationLabel}
                      </span>
                    </div>
                    <div style={{ flex: 1, minWidth: 0, position: 'relative', height: `${ROW_H}px` }}>
                      {(() => {
                        if (!t.start || !t.end) {
                          // Dates non exploitables ni calculables (ni date_effet+durée, ni date_fin+durée)
                          // — on affiche quand même le bail, ancré au début de la frise, plutôt que de le faire disparaître.
                          return (
                            <div
                              onMouseMove={e => setTooltip({ x: e.clientX, y: e.clientY, tenant: t })}
                              onMouseLeave={() => setTooltip(null)}
                              onClick={() => t.row && window.dispatchEvent(new CustomEvent('etatlocatif-select', { detail: t.row }))}
                              style={{
                                position: 'absolute', left: '6px', width: '220px', top: '14px', height: '24px', borderRadius: '5px',
                                display: 'flex', alignItems: 'center', padding: '0 10px', fontSize: '11px', fontWeight: 600,
                                color: 'var(--text3)', cursor: t.row ? 'pointer' : 'default',
                                background: 'repeating-linear-gradient(45deg, var(--surface2), var(--surface2) 5px, var(--border) 5px, var(--border) 10px)',
                              }}>
                              ⚠ Dates non déterminées
                            </div>
                          )
                        }
                        const left = ((t.start - domainStart) / domainMs) * 100
                        const segments = tenantSegments(t)
                        const barSpan = t.end - t.start
                        const fullEnd = effectiveEnd(t)
                        const fullSpan = fullEnd - t.start
                        const fullWidth = ((fullEnd - domainStart) / domainMs) * 100 - left
                        const fixedPortionPct = t.reconductionTacite ? (barSpan / fullSpan) * 100 : 100
                        return (
                          <div
                            onMouseMove={e => setTooltip({ x: e.clientX, y: e.clientY, tenant: t })}
                            onMouseLeave={() => setTooltip(null)}
                            onClick={() => t.row && window.dispatchEvent(new CustomEvent('etatlocatif-select', { detail: t.row }))}
                            style={{ position: 'absolute', left: `${left}%`, width: `${fullWidth}%`, top: '15px', height: '22px', cursor: t.row ? 'pointer' : 'default' }}>
                            <div style={{ position: 'relative', height: '100%', borderRadius: '5px', overflow: 'hidden', display: 'flex', border: t.estimated ? '1.5px dashed var(--text3)' : 'none' }}>
                              <div style={{ position: 'absolute', left: 0, width: `${fixedPortionPct}%`, top: 0, bottom: 0, display: 'flex' }}>
                                {segments.map((seg, si) => (
                                  <div key={si} style={{
                                    position: 'absolute', left: `${((seg.start - t.start) / barSpan) * 100}%`, width: `${((seg.end - seg.start) / barSpan) * 100}%`,
                                    top: 0, bottom: 0, background: seg.color, opacity: t.estimated ? 0.65 : 1, borderRight: si < segments.length - 1 ? '1.5px solid var(--surface)' : 'none',
                                  }} />
                                ))}
                              </div>
                              {t.reconductionTacite && (
                                <div title={t.reconductionTacite.dateLimiteAbsolue
                                  ? `Reconduction tacite — fin automatique et sans formalité le ${fmt(t.reconductionTacite.dateLimiteAbsolue)}`
                                  : 'Reconduction tacite — durée non figée, se poursuit sauf congé'} style={{
                                  position: 'absolute', left: `${fixedPortionPct}%`, right: 0, top: 0, bottom: 0,
                                  background: `repeating-linear-gradient(45deg, ${segments[segments.length-1]?.color || 'var(--accent)'}33, ${segments[segments.length-1]?.color || 'var(--accent)'}33 4px, transparent 4px, transparent 8px)`,
                                  borderLeft: '1.5px dashed var(--accent)',
                                  borderRight: t.reconductionTacite.dateLimiteAbsolue ? '3px solid var(--danger)' : 'none',
                                }} />
                              )}
                            </div>
                          </div>
                        )
                      })()}
                      <div title={`Aujourd'hui : ${fmt(today)}`} style={{
                        position: 'absolute', left: `${((today - domainStart) / domainMs) * 100}%`, top: 0, bottom: 0,
                        width: 0, borderLeft: '1.5px dashed var(--accent)', opacity: 0.55, zIndex: 2,
                      }} />
                    </div>
                  </div>
              )
              })}
            </div>
            )
          })()}
        </div>
      </div>
      {tooltip && <TenantTooltip x={tooltip.x} y={tooltip.y} tenant={tooltip.tenant} />}
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

  // Si un avenant confirme la date d'effet (ex: levée d'une condition
  // suspensive), on l'utilise comme référence — et on recalcule date_fin et
  // break_options à partir d'elle, sauf si un avenant a lui-même déjà fourni
  // une valeur explicite pour ces champs (auquel cas elle prime).
  let effetConfirmePar = null
  if (!isAv && Array.isArray(item.avenants) && item.avenants.length > 0) {
    const toSortableAv = s => { const m = String(s || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/); return m ? `${m[3]}-${m[2]}-${m[1]}` : String(s || '') }
    const sortedAvs = [...item.avenants].sort((a, b) =>
      toSortableAv(a.data?.date_effet_avenant || a.data?.date_signature_avenant || a.created_at)
        .localeCompare(toSortableAv(b.data?.date_effet_avenant || b.data?.date_signature_avenant || b.created_at)))
    let hasExplicitDateFin = false
    let hasExplicitBreaks = false
    sortedAvs.forEach(av => {
      const mods = av.data?.champs_modifies || {}
      if (mods.date_effet) { d.date_effet = mods.date_effet; effetConfirmePar = av }
      if (mods.date_fin) hasExplicitDateFin = true
      if (Array.isArray(mods.break_options) && mods.break_options.length > 0) hasExplicitBreaks = true
    })
    if (effetConfirmePar) {
      const startConfirmed = parseFrDate(d.date_effet)
      if (startConfirmed) {
        if (!hasExplicitDateFin) {
          const m = String(d.duree_totale || '').match(/(\d+)\s*ans?/i)
          if (m) {
            const end = new Date(startConfirmed.getFullYear() + parseInt(m[1]), startConfirmed.getMonth(), startConfirmed.getDate() - 1)
            d.date_fin = fmtFR(end)
          }
        }
        if (!hasExplicitBreaks) {
          const recalculed = computeBreaks(d.date_effet, d.date_fin, d.conditions_break, [], d.duree_ferme)
          if (recalculed.length > 0) d.break_options = recalculed
        }
      }
    }
  }

  // Enrichir les breaks à l'affichage aussi (données déjà en base non recalculées)
  const src = d._sources || {}
  const pages = item.data?._pages || {}
  // sanitizeBreakDates : filet de sécurité pour les données déjà en base où
  // break_options contiendrait une phrase descriptive au lieu d'une date pure.
  // computeBreaks tourne TOUJOURS et vient compléter (union) ce qui est déjà
  // stocké — pas seulement en fallback si vide — pour rattraper à l'affichage
  // les extractions faites avant le renforcement de ce calcul (ex: 3e break
  // triennale manquante alors que les 2 premières étaient déjà en base).
  let breaks = sanitizeBreakDates(d.break_options || [])
  const computedBreaks = computeBreaks(d.date_effet, d.date_fin, d.conditions_break, [], d.duree_ferme)
  if (computedBreaks.length > 0) {
    const existingSet = new Set(breaks.map(b => b.trim()))
    computedBreaks.forEach(c => { if (!existingSet.has(c)) { breaks.push(c); existingSet.add(c) } })
    breaks.sort((a, b) => { const da = parseFR(a), db = parseFR(b); return (da && db) ? da - db : 0 })
  }

  // Si duree_ferme est renseignée, supprimer les breaks AVANT date_effet + duree_ferme
  if (d.duree_ferme && d.date_effet) {
    const effet = parseFR(d.date_effet)
    const dfm = String(d.duree_ferme)
    const ymatch = dfm.match(/(\d+)\s*ans?/), mmatch = dfm.match(/(\d+)\s*mois/)
    const years = ymatch ? parseInt(ymatch[1]) : 0, months = mmatch ? parseInt(mmatch[1]) : 0
    if (effet && (years > 0 || months > 0)) {
      const minBreak = new Date(effet.getFullYear() + years, effet.getMonth() + months, effet.getDate() - 1)
      breaks = breaks.filter(b => {
        const bd = parseFR(b)
        return bd && bd >= minBreak
      })
      // Dédoublonner (ex: 29/06/2031 et 30/06/2031 = même date à 1 jour près)
      breaks = [...new Map(breaks.map(b => {
        const bd = parseFR(b)
        return [bd ? `${bd.getFullYear()}-${bd.getMonth()}` : b, b]
      })).values()]
    }
  }

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

  // Breaks conditionnelles (indemnites_break avec une break_date qui n'est PAS
  // déjà dans break_options) — jusqu'ici visibles seulement dans le texte libre
  // "Détail échéances", on les fait remonter dans la frise principale avec leur
  // condition, plutôt que de les laisser noyées dans un paragraphe.
  const cleanBreakParsedDates = breaks.map(b => parseFR(b)).filter(Boolean)
  const condBreakEntries = (Array.isArray(d.indemnites_break) ? d.indemnites_break : [])
    .map(ib => ({ date: ib.break_date ? normalizeDate(safeStr(ib.break_date)) : null, condition: safeStr(ib.motif) || safeStr(ib.calcul) }))
    // Filet de sécurité : une clause de CESSION (remboursement si cession du
    // bail/fonds à un tiers hors groupe) n'est pas une vraie option de sortie
    // du preneur — elle ne doit jamais apparaître comme une "break conditionnelle",
    // même si l'IA lui a par erreur attribué une break_date (cas d'exclusion déjà
    // prévu au prompt, mais pas toujours respecté).
    .filter(cb => !/cession/i.test(cb.condition || ''))
    .filter(cb => cb.date && /^\d{2}\/\d{2}\/\d{4}$/.test(cb.date))
    .filter(cb => {
      // Exclure si une break "propre" tombe à quelques jours près de cette
      // date — c'est très probablement LA MÊME échéance triennale, juste
      // calculée deux fois par l'IA (deux appels distincts, conventions de
      // jour légèrement différentes), pas une échéance véritablement en plus.
      const cbParsed = parseFR(cb.date)
      if (!cbParsed) return false
      const isDuplicateOfClean = cleanBreakParsedDates.some(bd => Math.abs(bd - cbParsed) <= 3 * 24 * 60 * 60 * 1000)
      return !isDuplicateOfClean
    })

  // Pour chaque break "propre", rattacher la condition/indemnité éventuellement
  // décrite dans indemnites_break pour la MÊME échéance (à quelques jours près)
  // — plutôt que de la perdre simplement parce qu'elle décrit une date déjà
  // affichée par ailleurs (cf. le filtre condBreakEntries ci-dessus).
  const rawIndemnitesBreak = (Array.isArray(d.indemnites_break) ? d.indemnites_break : [])
    .filter(ib => !/cession/i.test(safeStr(ib.motif) || '') && !/cession/i.test(safeStr(ib.calcul) || ''))
  function findIndemniteForBreak(breakDateStr) {
    const bd = parseFR(breakDateStr)
    if (!bd) return null
    const match = rawIndemnitesBreak.find(ib => {
      const ibDate = ib.break_date ? parseFR(normalizeDate(safeStr(ib.break_date))) : null
      return ibDate && Math.abs(ibDate - bd) <= 3 * 24 * 60 * 60 * 1000
    })
    return match ? (safeStr(match.motif) || safeStr(match.calcul)) : null
  }

  const breakItems = [
    ...breaks.map(br => ({ val: br, conditional: false, indemnite: findIndemniteForBreak(br) })),
    ...condBreakEntries.map(cb => ({ val: cb.date, conditional: true, condition: cb.condition })),
  ].sort((a, b) => { const da = parseFR(a.val), db = parseFR(b.val); return (da && db) ? da - db : 0 })

  let bNum = 0
  const primaryDates = [
    d.date_effet ? { key: 'date_effet', label: "Prise d'effet", type: 'primary', confirmedByAvenant: effetConfirmePar } : null,
    ...breakItems.map((item, i) => {
      if (!item.conditional) bNum++
      return {
        key: `break_${i}`,
        label: item.conditional ? 'Break conditionnelle' : `Break option${breakItems.filter(x => !x.conditional).length > 1 ? ' ' + bNum : ''}`,
        val: item.val,
        type: item.conditional ? 'break_conditionnel' : 'break',
        condition: item.condition,
        indemnite: item.indemnite,
        bNum: item.conditional ? null : bNum,
      }
    }),
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
            {show('preneur') && <div className="party-card"><div className="party-role">Preneur</div><div className="party-name" style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>{shortPartyName(d.preneur) || <span style={{ color: 'var(--text3)', fontStyle: 'italic', fontWeight: 400 }}>Non renseigné</span>}<PageJumpIcon item={item} pages={pages} field="preneur" /></div></div>}
            {show('bailleur') && <div className="party-card"><div className="party-role">Bailleur</div><div className="party-name" style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>{shortPartyName(d.bailleur) || <span style={{ color: 'var(--text3)', fontStyle: 'italic', fontWeight: 400 }}>Non renseigné</span>}<PageJumpIcon item={item} pages={pages} field="bailleur" /></div></div>}
            {show('garant') && d.garant && <div className="party-card" style={{gridColumn:'1/-1'}}><div className="party-role">Garant / Caution</div><div className="party-name">{shortPartyName(d.garant)}</div></div>}
          </div>
        </div>
      )}

      {/* Contrat */}
      {(show('type_bail') || show('duree_totale') || show('duree_ferme')) && (
        <div className="sec">
          <div className="sec-hd"><div className="sec-label">Contrat et durée</div></div>
          <div className="gx">
            <Field label="Type de contrat" value={d.type_bail} />
            <Field label="Durée totale" value={d.duree_totale} item={item} pages={pages} pageField="duree_totale" />
            <Field label="Durée ferme" value={d.duree_ferme} item={item} pages={pages} pageField="duree_ferme" />
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
              {primaryDates.map(f => {
                const isCondBreak = f.type === 'break_conditionnel'
                const pageField = f.type === 'break' ? 'break_options' : (isCondBreak ? 'indemnites_break' : f.key)
                return (
                <div key={f.key} className={`date-card${f.type === 'break' ? ' date-card-break' : ''}`}
                  style={isCondBreak ? { background: '#FAEEDA', border: '1px solid #EF9F27' } : undefined}>
                  <div className="date-lbl">
                    {f.type === 'break' && <span className="break-tag">B{breakItems.filter(x => !x.conditional).length > 1 ? f.bNum : ''}</span>}
                    {isCondBreak && <span className="break-tag" style={{ background: '#EF9F27' }} title="Break conditionnelle — voir la condition ci-dessous">⚠</span>}
                    {' '}{f.label}
                  </div>
                  <div className={`date-val${(f.type === 'break' || isCondBreak) ? ' break' : ''}`} style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap', color: isCondBreak ? '#B8860B' : undefined }}>
                    {f.val || d[f.key]}
                    <PageJumpIcon item={item} pages={pages} field={pageField} />
                  </div>
                  {f.type === 'break' && d.notice && <div style={{ fontSize: '11px', color: 'var(--text3)', marginTop: '4px' }}>Préavis : {d.notice}</div>}
                  {f.type === 'break' && f.indemnite && <div style={{ fontSize: '11px', color: '#B8860B', marginTop: '4px' }}>Indemnité si exercée : {f.indemnite}</div>}
                  {isCondBreak && f.condition && <div style={{ fontSize: '11px', color: '#B8860B', marginTop: '4px' }}>Conditionnelle : {f.condition}</div>}
                  {f.confirmedByAvenant && (
                    <div style={{ fontSize: '11px', color: 'var(--accent)', marginTop: '4px' }}>
                      Confirmée par avenant{f.confirmedByAvenant.data?.date_signature_avenant ? ` du ${f.confirmedByAvenant.data.date_signature_avenant}` : ''} — échéance et fin de bail recalculées
                    </div>
                  )}
                </div>
              )})}
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
          {show('reconduction_tacite') && d.reconduction_tacite?.applicable && (
            <div style={{
              marginTop: '10px', display: 'flex', alignItems: 'center', gap: '8px',
              padding: '9px 12px', borderRadius: 'var(--r)', background: 'var(--accent-bg)', border: '1px solid rgba(26,95,168,.15)',
            }}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" strokeWidth="2" style={{ flexShrink: 0 }}>
                <polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/>
              </svg>
              <span style={{ fontSize: '12.5px', color: 'var(--text2)' }}>
                <strong style={{ color: 'var(--text)' }}>Reconduction tacite</strong> au-delà du {normalizeDate(d.date_fin) || 'terme ferme'} — le bail se poursuit
                {d.reconduction_tacite.periodicite ? ` ${d.reconduction_tacite.periodicite}` : ''} sauf congé donné
                {d.reconduction_tacite.preavis ? ` avec un préavis de ${d.reconduction_tacite.preavis}` : ''}.
              </span>
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
            {show('surface_totale_m2') && (() => {
              const detailSumHero = (Array.isArray(d.surfaces_detail) ? d.surfaces_detail : [])
                .filter(r => !(r.categorie || '').toLowerCase().includes('station'))
                .reduce((a, r) => a + (parseFloat(String(r.surface_m2 || '').replace(',', '.')) || 0), 0)
              const totalHero = parseFloat(String(d.surface_totale_m2 || '').replace(',', '.')) || 0
              const gapHero = totalHero > 0 && detailSumHero > 0 ? totalHero - detailSumHero : 0
              const showGapNote = gapHero > Math.max(1, totalHero * 0.01)
              return (
                <div className="field">
                  <div className="field-lbl">Surface totale louée</div>
                  <div className="field-val" style={{ fontSize: '17px', fontWeight: 700, letterSpacing: '-0.02em', display: 'flex', alignItems: 'center', gap: '6px' }}>
                    {d.surface_totale_m2 ? `${d.surface_totale_m2} m²` : '—'}
                    <PageJumpIcon item={item} pages={pages} field="surface_totale_m2" />
                  </div>
                  {showGapNote && (
                    <div style={{ fontSize: '10.5px', color: 'var(--text3)', fontStyle: 'italic', marginTop: '2px' }}>
                      dont {gapHero.toLocaleString('fr-FR', { maximumFractionDigits: 2 })} m² de quote-part parties communes non ventilée
                    </div>
                  )}
                </div>
              )
            })()}
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
                    <div style={{ fontSize: '11px', fontWeight: 700, color: 'var(--text3)', textTransform: 'uppercase', letterSpacing: '.06em', marginBottom: '8px' }}>Modification de surface</div>
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
                <div className="loyer-lbl" style={{ display: "flex", alignItems: "center", gap: "4px" }}>Loyer HT/HC annuel à la signature{src.loyer_signature_montant && <span title={src.loyer_signature_montant} style={{ cursor: "help", fontSize: "10px", color: "var(--text3)", border: "1px solid var(--border2)", borderRadius: "50%", width: "13px", height: "13px", display: "inline-flex", alignItems: "center", justifyContent: "center", fontWeight: 700 }}>i</span>}</div>
                <div style={{ display: 'flex', alignItems: 'baseline', gap: '12px', flexWrap: 'wrap' }}>
                  <div className="loyer-amount">{fmtEur(d.loyer_signature_montant) || d.loyer_signature_montant}</div>
                  <PageJumpIcon item={item} pages={pages} field="loyer_signature_montant" />
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
                <SurfaceTable surfaces={enriched} totalDeclared={d.surface_totale_m2} totalLoyerDeclared={d.loyer_signature_montant} />
              </div>
            )
          })()}
          <div className="gx" style={{ marginBottom: '8px' }}>

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


      {(() => {
        // Même filet de sécurité que pour la carte "Break conditionnelle" :
        // une clause de cession n'est pas une vraie indemnité de break, même
        // si l'IA l'a par erreur incluse dans indemnites_break.
        const cleanIndemnitesBreak = (d.indemnites_break || []).filter(row =>
          !/cession/i.test(safeStr(row.motif) || '') && !/cession/i.test(safeStr(row.calcul) || '')
        )
        if (cleanIndemnitesBreak.length === 0) return null
        return (
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
                  {cleanIndemnitesBreak.map((row, i) => (
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
        )
      })()}

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

function ActifPicker({ currentValue, existingGroups, onSave, onClose, anchorRect }) {
  const [q, setQ] = useState('')
  const inputRef = useRef()
  useEffect(() => { inputRef.current?.focus() }, [])

  const filtered = existingGroups.filter(g => g.toLowerCase().includes(q.toLowerCase()) && g !== currentValue)
  const showCreate = q.trim() && !existingGroups.map(g => g.toLowerCase()).includes(q.trim().toLowerCase())

  // Rendu en portail directement sur <body> : les conteneurs du tableau ont un
  // overflow (scroll) qui tronquerait un dropdown positionné en absolute normal.
  // Rendu en portail directement sur <body> : les conteneurs du tableau ont un
  // overflow (scroll) qui tronquerait un dropdown positionné en absolute normal.
  // Bascule au-dessus du bouton si pas assez de place en dessous (bas d'écran).
  const ESTIMATED_HEIGHT = 260 // input + liste, approximatif mais suffisant pour décider
  const spaceBelow = anchorRect ? window.innerHeight - anchorRect.bottom : Infinity
  const openAbove = anchorRect && spaceBelow < ESTIMATED_HEIGHT

  const style = anchorRect
    ? openAbove
      ? { position: 'fixed', bottom: window.innerHeight - anchorRect.top + 4, left: anchorRect.left, maxHeight: `${anchorRect.top - 8}px` }
      : { position: 'fixed', top: anchorRect.bottom + 4, left: anchorRect.left, maxHeight: `${spaceBelow - 8}px` }
    : { position: 'absolute', top: '100%', left: 0, marginTop: '2px' }

  return createPortal(
    <div style={{ ...style,
      background: 'var(--surface)', border: '1px solid var(--border2)', borderRadius: '8px',
      boxShadow: '0 8px 24px rgba(0,0,0,.18)', width: '220px', overflow: 'hidden', zIndex: 9999,
      display: 'flex', flexDirection: 'column' }}
      onClick={e => e.stopPropagation()}>
      <div style={{ padding: '6px', flexShrink: 0 }}>
        <input
          ref={inputRef}
          value={q}
          onChange={e => setQ(e.target.value)}
          onKeyDown={e => {
            if (e.key === 'Escape') onClose()
            if (e.key === 'Enter' && q.trim()) onSave(q.trim())
          }}
          placeholder="Rechercher ou créer…"
          style={{ width: '100%', boxSizing: 'border-box', padding: '5px 8px', fontSize: '12px',
            border: '1px solid var(--border2)', borderRadius: '5px', outline: 'none',
            background: 'var(--surface2)', color: 'var(--text)' }}
        />
      </div>
      <div style={{ overflowY: 'auto', minHeight: 0, flex: '1 1 auto' }}>
        {currentValue && (
          <div onClick={() => onSave('')}
            style={{ padding: '6px 12px', fontSize: '12px', cursor: 'pointer', color: 'var(--danger)',
              borderTop: '1px solid var(--border)' }}>
            ✕ Retirer du groupe
          </div>
        )}
        {filtered.map(g => (
          <div key={g} onClick={() => onSave(g)}
            style={{ padding: '6px 12px', fontSize: '12px', cursor: 'pointer', color: 'var(--text)',
              borderTop: '1px solid var(--border)' }}
            onMouseEnter={e => e.currentTarget.style.background = 'var(--accent-bg)'}
            onMouseLeave={e => e.currentTarget.style.background = ''}>
            📁 {g}
          </div>
        ))}
        {showCreate && (
          <div onClick={() => onSave(q.trim())}
            style={{ padding: '6px 12px', fontSize: '12px', cursor: 'pointer', color: 'var(--accent)',
              fontWeight: 600, borderTop: '1px solid var(--border)' }}
            onMouseEnter={e => e.currentTarget.style.background = 'var(--accent-bg)'}
            onMouseLeave={e => e.currentTarget.style.background = ''}>
            + Créer "{q.trim()}"
          </div>
        )}
        {!filtered.length && !showCreate && !currentValue && (
          <div style={{ padding: '8px 12px', fontSize: '12px', color: 'var(--text3)', fontStyle: 'italic' }}>
            Aucun groupe existant
          </div>
        )}
      </div>
    </div>,
    document.body
  )
}

// ─── Modale de contrôle qualité ──────────────────────────────────────────────
function QualityCheckModal({ bails, onClose, onSelect, onDismiss, onFixAnniversary }) {
  const [showDismissed, setShowDismissed] = useState(false)
  const [pending, setPending] = useState({}) // { [rowId]: true } — évite double-clic pendant l'écriture
  const [fixPending, setFixPending] = useState({}) // { [rowId]: true } — pour le bouton "−1 jour"
  const allResults = useMemo(() => {
    const dupDetails = findDuplicateBails(bails)
    return bails.map(row => {
      const r = auditBail(row)
      if (dupDetails[row.id]) {
        r.issues = [...r.issues, { type: 'doublon_suspect', severity: 'high', detail: dupDetails[row.id] }]
      }
      return r
    }).filter(r => r.issues.length > 0)
  }, [bails])
  const activeResults = allResults.filter(r => !r.dismissed)
  const dismissedResults = allResults.filter(r => r.dismissed)
  const severityColor = { high: 'var(--danger)', medium: 'var(--accent)', low: 'var(--text3)' }
  const severityBg = { high: 'var(--danger-bg)', medium: 'var(--accent-bg)', low: 'var(--surface2)' }
  const severityLabel = { high: 'À vérifier en priorité', medium: 'À vérifier', low: 'Info' }

  async function handleDismiss(e, rowId, dismissed) {
    e.stopPropagation()
    if (pending[rowId]) return
    setPending(prev => ({ ...prev, [rowId]: true }))
    await onDismiss?.(rowId, dismissed)
    setPending(prev => { const n = { ...prev }; delete n[rowId]; return n })
  }

  async function handleFixAnniversary(e, rowId, dates) {
    e.stopPropagation()
    if (fixPending[rowId]) return
    setFixPending(prev => ({ ...prev, [rowId]: true }))
    await onFixAnniversary?.(rowId, dates)
    setFixPending(prev => { const n = { ...prev }; delete n[rowId]; return n })
  }

  function ResultCard({ r, dismissedCard }) {
    return (
      <div key={r.row.id} onClick={() => onSelect(r.row)}
        style={{ border: '1px solid var(--border)', borderRadius: '10px', padding: '12px 14px', cursor: 'pointer', opacity: dismissedCard ? 0.6 : 1 }}
        onMouseEnter={e => e.currentTarget.style.background = 'var(--surface2)'}
        onMouseLeave={e => e.currentTarget.style.background = ''}>
        <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: '10px', marginBottom: '6px' }}>
          <div style={{ fontWeight: 700, fontSize: '13px' }}>{r.label}</div>
          <button
            onClick={e => handleDismiss(e, r.row.id, !dismissedCard)}
            disabled={pending[r.row.id]}
            title={dismissedCard ? 'Réafficher dans la liste active' : 'Marquer comme vérifié — ne réapparaîtra plus'}
            style={{
              flexShrink: 0, fontSize: '11px', fontWeight: 600, padding: '3px 9px', borderRadius: '999px', cursor: pending[r.row.id] ? 'default' : 'pointer',
              border: `1px solid ${dismissedCard ? 'var(--border2)' : 'var(--success)'}`,
              background: dismissedCard ? 'var(--surface)' : 'var(--success-bg)',
              color: dismissedCard ? 'var(--text3)' : 'var(--success)',
              opacity: pending[r.row.id] ? 0.5 : 1,
            }}>
            {dismissedCard ? '↺ Réafficher' : '✓ Vérifié'}
          </button>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: '5px' }}>
          {r.issues.map((iss, i) => (
            <div key={i} style={{ display: 'flex', alignItems: 'flex-start', gap: '8px', flexWrap: 'wrap' }}>
              <span style={{
                fontSize: '10px', fontWeight: 700, padding: '1px 7px', borderRadius: '999px', flexShrink: 0, marginTop: '1px',
                color: severityColor[iss.severity], background: severityBg[iss.severity],
              }}>{severityLabel[iss.severity]}</span>
              <span style={{ fontSize: '12px', color: 'var(--text2)', lineHeight: 1.4 }}>{iss.detail}</span>
              {iss.type === 'break_jour_anniversaire' && !dismissedCard && (
                <button
                  onClick={e => handleFixAnniversary(e, r.row.id, iss.affectedDates)}
                  disabled={fixPending[r.row.id]}
                  title="Retirer 1 jour à cette/ces date(s) — corrige uniquement les breaks concernées, ne touche à rien d'autre"
                  style={{
                    fontSize: '10.5px', fontWeight: 600, padding: '2px 8px', borderRadius: '999px', cursor: fixPending[r.row.id] ? 'default' : 'pointer',
                    border: '1px solid var(--accent)', background: 'var(--accent-bg)', color: 'var(--accent)',
                    opacity: fixPending[r.row.id] ? 0.5 : 1, flexShrink: 0,
                  }}>
                  {fixPending[r.row.id] ? '…' : '−1 jour'}
                </button>
              )}
            </div>
          ))}
        </div>
      </div>
    )
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" style={{ width: '820px', maxHeight: '85vh' }} onClick={e => e.stopPropagation()}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div>
            <div className="modal-title">Contrôle qualité</div>
            <div style={{ fontSize: '12px', color: 'var(--text3)', marginTop: '2px' }}>
              Détection heuristique — ne modifie rien, à vérifier/réextraire manuellement au cas par cas
            </div>
          </div>
          <button onClick={onClose} title="Fermer" style={{ background: 'none', border: 'none', fontSize: '20px', lineHeight: 1, cursor: 'pointer', color: 'var(--text2)', padding: '4px' }}>✕</button>
        </div>

        <div style={{ overflowY: 'auto', flex: 1, paddingRight: '4px' }}>
          {allResults.length === 0 ? (
            <div style={{ padding: '40px', textAlign: 'center', color: 'var(--text3)', fontSize: '13px' }}>
              Aucun cas suspect détecté sur les {bails.length} bail{bails.length !== 1 ? 'x' : ''} analysé{bails.length !== 1 ? 's' : ''}.
            </div>
          ) : (
            <>
              <div style={{ fontSize: '12px', color: 'var(--text3)', marginBottom: '10px' }}>
                {activeResults.length} bail{activeResults.length !== 1 ? 'x' : ''} sur {bails.length} présentent au moins un point à vérifier.
                {dismissedResults.length > 0 && (
                  <span onClick={() => setShowDismissed(v => !v)} style={{ marginLeft: '8px', color: 'var(--accent)', cursor: 'pointer', textDecoration: 'underline' }}>
                    {showDismissed ? 'Masquer' : 'Afficher aussi'} les {dismissedResults.length} déjà vérifié{dismissedResults.length !== 1 ? 's' : ''}
                  </span>
                )}
              </div>
              {activeResults.length === 0 ? (
                <div style={{ padding: '24px', textAlign: 'center', color: 'var(--text3)', fontSize: '13px' }}>
                  Tout est vérifié pour l'instant 🎉
                </div>
              ) : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  {activeResults.map(r => <ResultCard key={r.row.id} r={r} dismissedCard={false} />)}
                </div>
              )}
              {showDismissed && dismissedResults.length > 0 && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginTop: '14px', paddingTop: '14px', borderTop: '1px solid var(--border)' }}>
                  {dismissedResults.map(r => <ResultCard key={r.row.id} r={r} dismissedCard={true} />)}
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  )
}

// ─── Modale d'attache en masse (dossier → correspondance par nom de fichier) ─
function BulkAttachModal({ candidateRows, allRows, onClose, onRefresh }) {
  const [dragging, setDragging] = useState(false)
  const [scanning, setScanning] = useState(false)
  const [matches, setMatches] = useState(null) // null tant qu'aucun dossier déposé
  const [progress, setProgress] = useState(null) // { current, total, fileName, state }
  const [results, setResults] = useState(null) // { success, failed: [{name, msg}] }
  const inputRef = useRef()

  function buildMatches(files) {
    const usedIds = new Set()
    return files.map(file => {
      const norm = file.name.trim().toLowerCase()
      const row = candidateRows.find(r => !usedIds.has(r.id) && (r.file_name || '').trim().toLowerCase() === norm)
      if (row) { usedIds.add(row.id); return { file, row, status: 'matched' } }
      // Pas de candidat à attacher — mais peut-être qu'un document du même nom
      // existe déjà et a simplement déjà son fichier source (rien à faire),
      // plutôt qu'aucun document correspondant n'existe du tout en base.
      const alreadyAttached = (allRows || []).find(r => r.storage_path && (r.file_name || '').trim().toLowerCase() === norm)
      if (alreadyAttached) return { file, row: null, status: 'already_attached' }
      return { file, row: null, status: 'no_match' }
    })
  }

  const handleFiles = useCallback(files => {
    const valid = Array.from(files).filter(f => ['pdf', 'docx'].includes(f.name.split('.').pop().toLowerCase()))
    if (!valid.length) { alert('Aucun fichier PDF ou DOCX trouvé.'); return }
    setMatches(buildMatches(valid))
  }, [candidateRows])

  const handleDrop = useCallback(async e => {
    e.preventDefault()
    setDragging(false)
    const items = Array.from(e.dataTransfer.items || [])
    const hasEntries = items.some(i => i.webkitGetAsEntry)
    if (!hasEntries) { handleFiles(e.dataTransfer.files); return }
    setScanning(true)
    const entries = items.map(i => i.webkitGetAsEntry()).filter(Boolean)
    const allFiles = (await Promise.all(entries.map(en => collectFiles(en, '')))).flat()
    setScanning(false)
    if (!allFiles.length) { alert('Aucun fichier PDF ou DOCX trouvé dans le répertoire.'); return }
    handleFiles(allFiles)
  }, [handleFiles])

  async function runBulk() {
    const toRun = matches.filter(m => m.row)
    let success = 0
    const failed = []
    for (let i = 0; i < toRun.length; i++) {
      const { file, row } = toRun[i]
      const isAv = row.document_type === 'avenant'
      setProgress({ current: i + 1, total: toRun.length, fileName: file.name, state: 'compressing' })
      try {
        const prepared = await compressPdfIfNeeded(file, (c, t) => {
          setProgress(p => ({ ...p, state: 'compressing', progCurrent: c, progTotal: t }))
        })
        if (prepared.size > 30 * 1024 * 1024) {
          throw new Error(`Fichier trop volumineux (${Math.round(prepared.size / 1024 / 1024)} Mo > 30 Mo)`)
        }
        setProgress(p => ({ ...p, state: 'loading' }))
        const base64 = await toBase64(prepared)
        const mediaType = getMediaType(prepared)
        const extracted = await callClaude(base64, mediaType, isAv ? AVENANT_PROMPT : EXTRACTION_PROMPT)

        try {
          if (!isAv) {
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
          } else {
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
          }
        } catch (_) { /* non bloquant */ }

        await uploadSourceFile(row.id, prepared)
        const { error } = await supabase.from('extractions').update({ data: extracted }).eq('id', row.id)
        if (error) throw error
        success++
      } catch (err) {
        failed.push({ name: file.name, msg: err.message || 'Erreur inconnue' })
      }
    }
    setProgress(null)
    setResults({ success, failed })
    onRefresh?.()
  }

  const matchedCount = matches ? matches.filter(m => m.status === 'matched').length : 0
  const alreadyAttachedCount = matches ? matches.filter(m => m.status === 'already_attached').length : 0
  const noMatchCount = matches ? matches.filter(m => m.status === 'no_match').length : 0

  return (
    <div className="modal-overlay" onClick={() => { if (!progress) onClose() }}>
      <div className="modal" style={{ width: '720px', maxHeight: '85vh' }} onClick={e => e.stopPropagation()}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div>
            <div className="modal-title">Attacher un dossier en masse</div>
            <div style={{ fontSize: '12px', color: 'var(--text3)', marginTop: '2px' }}>
              Rapproche les fichiers du dossier aux documents déjà extraits mais sans fichier source, par nom de fichier identique, puis relance l'extraction pour chacun avec les règles actuelles.
            </div>
          </div>
          {!progress && <button onClick={onClose} title="Fermer" style={{ background: 'none', border: 'none', fontSize: '20px', lineHeight: 1, cursor: 'pointer', color: 'var(--text2)', padding: '4px' }}>✕</button>}
        </div>

        <div style={{ overflowY: 'auto', flex: 1, paddingRight: '4px' }}>
          {progress ? (
            <div style={{ padding: '30px 10px', textAlign: 'center' }}>
              <div style={{
                width: '36px', height: '36px', margin: '0 auto 16px', borderRadius: '50%',
                border: '3px solid var(--border2)', borderTopColor: 'var(--accent)',
                animation: 'spin 0.8s linear infinite',
              }} />
              <div style={{ fontWeight: 700, fontSize: '15px', marginBottom: '6px' }}>
                Traitement {progress.current}/{progress.total}
              </div>
              <div style={{ fontSize: '13px', color: 'var(--text2)', marginBottom: '10px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                {progress.fileName}
              </div>
              <div style={{ fontSize: '12px', color: 'var(--text3)', marginBottom: '14px' }}>
                {progress.state === 'compressing'
                  ? `Compression${progress.progTotal ? ` (page ${progress.progCurrent}/${progress.progTotal})` : '…'}`
                  : 'Extraction en cours…'}
              </div>
              <div className="progress-track" style={{ margin: '0 40px' }}><div className="progress-bar active" /></div>
            </div>
          ) : results ? (
            <div style={{ padding: '10px 4px' }}>
              <div style={{ fontWeight: 700, fontSize: '15px', marginBottom: '10px' }}>
                {results.success > 0 && <span style={{ color: 'var(--success)' }}>✓ {results.success} document{results.success > 1 ? 's' : ''} attaché{results.success > 1 ? 's' : ''} et réextrait{results.success > 1 ? 's' : ''}</span>}
                {results.success > 0 && results.failed.length > 0 && ' · '}
                {results.failed.length > 0 && <span style={{ color: 'var(--danger)' }}>✕ {results.failed.length} échec{results.failed.length > 1 ? 's' : ''}</span>}
              </div>
              {results.failed.length > 0 && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', marginBottom: '16px' }}>
                  {results.failed.map((f, i) => (
                    <div key={i} style={{ fontSize: '12px', padding: '8px 12px', background: 'var(--danger-bg)', borderRadius: '6px' }}>
                      <strong>{f.name}</strong> — {f.msg}
                    </div>
                  ))}
                </div>
              )}
              <button className="btn primary" onClick={onClose} style={{ width: '100%' }}>Fermer</button>
            </div>
          ) : matches ? (
            <div>
              <div style={{ fontSize: '13px', marginBottom: '10px' }}>
                <strong>{matchedCount}</strong> fichier{matchedCount > 1 ? 's' : ''} rapproché{matchedCount > 1 ? 's' : ''} avec succès
                {alreadyAttachedCount > 0 && <span style={{ color: 'var(--text3)' }}> · {alreadyAttachedCount} déjà attaché{alreadyAttachedCount > 1 ? 's' : ''} (rien à faire)</span>}
                {noMatchCount > 0 && <span style={{ color: 'var(--text3)' }}> · {noMatchCount} sans document correspondant en base</span>}
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', maxHeight: '380px', overflowY: 'auto', marginBottom: '16px' }}>
                {matches.map((m, i) => (
                  <div key={i} style={{
                    display: 'flex', alignItems: 'center', gap: '10px', padding: '8px 12px', borderRadius: '6px',
                    background: m.status === 'matched' ? 'var(--success-bg)' : m.status === 'already_attached' ? 'var(--accent-bg)' : 'var(--surface2)',
                  }}>
                    <span style={{ fontSize: '13px', flexShrink: 0 }}>{m.status === 'matched' ? '✓' : m.status === 'already_attached' ? '📄' : '—'}</span>
                    <span style={{ fontSize: '12px', flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{m.file.name}</span>
                    {m.status === 'matched' ? (
                      <span style={{ fontSize: '11px', color: 'var(--text3)', flexShrink: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '220px' }}>
                        → {m.row.data?.immeuble || m.row.data?.adresse || m.row.file_name}
                      </span>
                    ) : m.status === 'already_attached' ? (
                      <span style={{ fontSize: '11px', color: 'var(--accent)', flexShrink: 0, fontStyle: 'italic' }}>déjà attaché — rien à faire</span>
                    ) : (
                      <span style={{ fontSize: '11px', color: 'var(--text3)', flexShrink: 0, fontStyle: 'italic' }}>aucun document correspondant en base</span>
                    )}
                  </div>
                ))}
              </div>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button className="btn primary" onClick={runBulk} disabled={matchedCount === 0} style={{ flex: 1, opacity: matchedCount === 0 ? 0.5 : 1 }}>
                  Attacher et réextraire {matchedCount > 0 ? `(${matchedCount})` : ''}
                </button>
                <button className="btn" onClick={() => setMatches(null)}>Recommencer</button>
              </div>
            </div>
          ) : (
            <div
              className={`drop-zone${dragging ? ' dragging' : ''}`}
              onClick={() => inputRef.current?.click()}
              onDragOver={e => { e.preventDefault(); setDragging(true) }}
              onDragLeave={e => { if (!e.currentTarget.contains(e.relatedTarget)) setDragging(false) }}
              onDrop={handleDrop}
            >
              <input ref={inputRef} type="file" accept=".pdf,.docx" multiple style={{ display: 'none' }} onChange={e => handleFiles(e.target.files)} />
              <div className="drop-icon">
                <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.2">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                  <polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>
                </svg>
              </div>
              <div className="drop-title">
                {scanning ? '⏳ Scan du répertoire en cours…' : 'Déposez un dossier ou des fichiers ici'}
              </div>
              <div className="drop-sub">
                {candidateRows.length} document{candidateRows.length > 1 ? 's' : ''} sans fichier source actuellement — le rapprochement se fait par nom de fichier identique
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

function Dashboard({ tree, totalCounts, onSelect, onDelete, onClear, onExportAll, newIds, onRefresh, onUpdateActif, onNewAvenant, filter, setFilter, search, setSearch }) {
  const [confirmClear, setConfirmClear] = useState(false)
  const [exportErrors, setExportErrors] = useState(null)
  const [extractionErrors, setExtractionErrors] = useState(null) // null or array of {name, reason}
  const [confirmDelete, setConfirmDelete] = useState(null) // item to delete
  const [avenantTarget, setAvenantTarget] = useState(null) // bail row en attente d'ajout d'avenant
  const [avenantUpload, setAvenantUpload] = useState({}) // { [bailId]: { state: 'compressing'|'loading'|'error', error, progress } } — utilisé par l'attache de fichier source uniquement
  const [avenantBatchProgress, setAvenantBatchProgress] = useState(null) // { bailLabel, current, total, fileName, state } — bloquant, pour l'ajout d'avenant(s)
  const [attachTarget, setAttachTarget] = useState(null) // document en attente d'attache de fichier source
  const [confirmReextract, setConfirmReextract] = useState(null) // row en attente de confirmation de réextraction
  const [confirmAttachReextract, setConfirmAttachReextract] = useState(null) // row sans fichier source, en attente de confirmation attache+réextraction
  const [showBulkAttach, setShowBulkAttach] = useState(false)
  const [reextractProgress, setReextractProgress] = useState(null) // { label, state } — bloquant
  const [toast, setToast] = useState(null) // { type: 'success'|'error', message }
  const avenantInputRef = useRef(null)
  const attachInputRef = useRef(null)
  const toastTimerRef = useRef(null)

  function showToast(type, message) {
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current)
    setToast({ type, message })
    toastTimerRef.current = setTimeout(() => setToast(null), 4500)
  }

  // Flatten all items for table
  const [expanded, setExpanded] = useState({})
  const [sortDir, setSortDir] = useState('asc')
  const [sortBy, setSortBy] = useState('actif') // 'actif' | 'preneur'
  const [editingActif, setEditingActif] = useState(null) // bail id
  const [editingActifRect, setEditingActifRect] = useState(null) // position du bouton cliqué
  const [renamingGroup, setRenamingGroup] = useState(null) // group name

  // Close picker on outside click
  useEffect(() => {
    if (!editingActif) return
    const handler = () => setEditingActif(null)
    document.addEventListener('click', handler)
    return () => document.removeEventListener('click', handler)
  }, [editingActif])

  // Derive all existing actif groups from tree
  const existingGroups = [...new Set(tree.map(b => b.actif_group).filter(Boolean))].sort()

  async function renameGroup(oldName, newName) {
    const v = newName.trim()
    if (!v || v === oldName) { setRenamingGroup(null); return }
    const ids = tree.filter(b => b.actif_group === oldName).map(b => b.id)
    for (const id of ids) {
      await supabase.from('extractions').update({ actif_group: v }).eq('id', id)
      await supabase.from('extractions').update({ actif_group: v }).eq('parent_id', id)
      onUpdateActif?.(id, v)
    }
    setRenamingGroup(null)
  }

  const savingRef = useRef(false)

  async function saveActifGroup(id, value) {
    if (savingRef.current) return
    savingRef.current = true
    const v = (value || '').trim()
    setEditingActif(null)
    onUpdateActif?.(id, v)
    const { error } = await supabase.from('extractions').update({ actif_group: v || null }).eq('id', id)
    if (!error) await supabase.from('extractions').update({ actif_group: v || null }).eq('parent_id', id)
    savingRef.current = false
  }

  function toggleExpand(id) {
    setExpanded(prev => ({ ...prev, [id]: !prev[id] }))
  }

  // Ajout d'un avenant directement depuis le dashboard, sur un bail déjà extrait
  function openAvenantPicker(bailRow) {
    setAvenantTarget(bailRow)
    // Laisse le state se poser avant de déclencher le picker natif
    setTimeout(() => avenantInputRef.current?.click(), 0)
  }

  async function handleAvenantFile(e) {
    const fileList = Array.from(e.target.files || [])
    const bailRow = avenantTarget
    e.target.value = '' // reset l'input pour permettre de re-choisir les mêmes fichiers
    if (!fileList.length || !bailRow) return

    const validFiles = fileList.filter(f => ['pdf', 'docx'].includes(f.name.split('.').pop().toLowerCase()))
    if (!validFiles.length) {
      alert('Format non supporté. PDF ou DOCX uniquement.')
      return
    }

    const label = bailRow.data?.immeuble || bailRow.data?.adresse || bailRow.file_name
    let successCount = 0
    const failedFiles = []

    for (let idx = 0; idx < validFiles.length; idx++) {
      const file = validFiles[idx]
      setAvenantBatchProgress({ bailLabel: label, current: idx + 1, total: validFiles.length, fileName: file.name, state: 'compressing' })
      try {
        const prepared = await compressPdfIfNeeded(file, (current, total) => {
          setAvenantBatchProgress(prev => ({ ...prev, state: 'compressing', progCurrent: current, progTotal: total }))
        })

        if (prepared.size > 30 * 1024 * 1024) {
          throw new Error(`Fichier trop volumineux (${Math.round(prepared.size / 1024 / 1024)} Mo > 30 Mo) — compressez le PDF avant de déposer.`)
        }

        setAvenantBatchProgress(prev => ({ ...prev, state: 'loading' }))
        const base64 = await toBase64(prepared)
        const mediaType = getMediaType(prepared)
        const extracted = await callClaude(base64, mediaType, AVENANT_PROMPT)

        const { data: saved, error } = await supabase.from('extractions').insert({
          file_name: file.name,
          data: extracted,
          document_type: 'avenant',
          parent_id: bailRow.id,
          actif_group: bailRow.actif_group || null,
        }).select().single()
        if (error) throw error
        if (saved?.id) await uploadSourceFile(saved.id, prepared) // on attend la fin pour éviter un rafraîchissement prématuré du dashboard

        successCount++
        if (saved?.id) onNewAvenant?.(saved.id)
      } catch (err) {
        failedFiles.push({ name: file.name, msg: err.message || 'Erreur inconnue' })
      }
    }

    setAvenantBatchProgress(null)
    setAvenantTarget(null)
    setExpanded(prev => ({ ...prev, [bailRow.id]: true })) // déplie le bail pour montrer les nouveaux avenants

    if (successCount > 0) {
      showToast('success', `${successCount} avenant${successCount > 1 ? 's' : ''} ajouté${successCount > 1 ? 's' : ''} à « ${label} »`)
    }
    if (failedFiles.length > 0) {
      showToast('error', `${failedFiles.length} échec${failedFiles.length > 1 ? 's' : ''} : ${failedFiles.map(f => f.name).join(', ')}`)
    }
    onRefresh?.()
  }

  function viewSourceFile(row) {
    openSourceAtPage(row, 1)
  }

  // ─── Rattrapage : attacher le fichier source à un document déjà extrait ────
  function buildValuesToLocate(row) {
    const values = {}
    if (row.document_type === 'avenant') {
      const mods = row.data?.champs_modifies || {}
      const candidates = {
        objet_avenant: row.data?.objet_avenant,
        date_effet_avenant: row.data?.date_effet_avenant,
        date_signature_avenant: row.data?.date_signature_avenant,
        loyer_signature_montant: mods.loyer_signature_montant,
        duree_ferme: mods.duree_ferme,
        date_effet: mods.date_effet,
        date_fin: mods.date_fin,
        break_options: mods.break_options,
        surface_totale_m2: mods.surface_totale_m2,
        depot_garantie_montant: mods.depot_garantie_montant,
      }
      Object.entries(candidates).forEach(([k, v]) => { if (v != null && v !== '') values[k] = v })
    } else {
      const d = row.data || {}
      const candidates = {
        preneur: d.preneur, bailleur: d.bailleur,
        date_effet: d.date_effet, date_fin: d.date_fin,
        break_options: d.break_options,
        surface_totale_m2: d.surface_totale_m2,
        loyer_signature_montant: d.loyer_signature_montant,
      }
      Object.entries(candidates).forEach(([k, v]) => {
        if (v != null && v !== '' && !(Array.isArray(v) && v.length === 0)) values[k] = v
      })
    }
    return values
  }

  function openAttachPicker(row) {
    setAttachTarget(row)
    setTimeout(() => attachInputRef.current?.click(), 0)
  }

  async function handleAttachFile(e) {
    const file = e.target.files?.[0]
    const row = attachTarget
    e.target.value = ''
    if (!file || !row) return

    // Les baux sans fichier source sont typiquement d'anciennes extractions,
    // faites avec des règles de prompt potentiellement dépassées — on profite
    // donc de l'attache du fichier pour relancer une extraction complète dans
    // la foulée, plutôt que de se limiter à la localisation de pages.
    const isAv = row.document_type === 'avenant'
    const label = row.data?.immeuble || row.data?.adresse || row.file_name
    setReextractProgress({ label, state: 'compressing' })
    try {
      const prepared = await compressPdfIfNeeded(file, (current, total) => {
        setReextractProgress(prev => ({ ...prev, state: 'compressing', progCurrent: current, progTotal: total }))
      })
      if (prepared.size > 30 * 1024 * 1024) {
        throw new Error(`Fichier trop volumineux (${Math.round(prepared.size / 1024 / 1024)} Mo > 30 Mo)`)
      }

      setReextractProgress(prev => ({ ...prev, state: 'loading' }))
      const base64 = await toBase64(prepared)
      const mediaType = getMediaType(prepared)
      const extracted = await callClaude(base64, mediaType, isAv ? AVENANT_PROMPT : EXTRACTION_PROMPT)

      // Mêmes appels complémentaires (breaks + financier) qu'à l'extraction initiale.
      try {
        if (!isAv) {
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
        } else {
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
        }
      } catch (_) { /* non bloquant */ }

      await uploadSourceFile(row.id, prepared) // attache le fichier + renseigne storage_path
      const { error: updateErr } = await supabase.from('extractions').update({ data: extracted }).eq('id', row.id)
      if (updateErr) throw updateErr

      setReextractProgress(null)
      setAttachTarget(null)
      showToast('success', `« ${label} » attaché et réextrait avec succès`)
      onRefresh?.()
    } catch (err) {
      setReextractProgress(null)
      showToast('error', `Échec de l'attache/réextraction : ${err.message || 'Erreur inconnue'}`)
    }
  }

  // ─── Réextraction en place (garde l'id, ne touche pas aux avenants) ────────
  async function handleReextract(row) {
    setConfirmReextract(null)
    const isAv = row.document_type === 'avenant'
    const label = row.data?.immeuble || row.data?.adresse || row.file_name
    setReextractProgress({ label, state: 'downloading' })
    try {
      if (!row.storage_path) throw new Error('Aucun fichier source attaché à ce document.')
      const { data: signedData, error: signErr } = await supabase.storage
        .from('lease-sources').createSignedUrl(row.storage_path, 300)
      if (signErr) throw signErr
      const fileRes = await fetch(signedData.signedUrl)
      if (!fileRes.ok) throw new Error('Téléchargement du fichier source échoué')
      const blob = await fileRes.blob()
      let file = new File([blob], row.file_name || row.storage_path.split('/').pop(), { type: blob.type })

      setReextractProgress(prev => ({ ...prev, state: 'compressing' }))
      file = await compressPdfIfNeeded(file, (current, total) => {
        setReextractProgress(prev => ({ ...prev, state: 'compressing', progCurrent: current, progTotal: total }))
      })
      if (file.size > 30 * 1024 * 1024) {
        throw new Error(`Fichier trop volumineux (${Math.round(file.size / 1024 / 1024)} Mo > 30 Mo)`)
      }

      setReextractProgress(prev => ({ ...prev, state: 'loading' }))
      const base64 = await toBase64(file)
      const mediaType = getMediaType(file)
      const extracted = await callClaude(base64, mediaType, isAv ? AVENANT_PROMPT : EXTRACTION_PROMPT)

      // Appels dédiés en parallèle : breaks + financier — mêmes enrichissements
      // que lors d'une extraction initiale, pour ne pas produire un résultat
      // moins complet qu'une première extraction.
      try {
        if (!isAv) {
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
        } else {
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
        }
      } catch (_) { /* non bloquant */ }

      // On garde l'id, le parent_id, l'actif_group et le storage_path déjà en
      // place — seul le contenu extrait (data) est remplacé.
      const { error: updateErr } = await supabase.from('extractions').update({ data: extracted }).eq('id', row.id)
      if (updateErr) throw updateErr

      setReextractProgress(null)
      showToast('success', `« ${label} » réextrait avec succès`)
      onRefresh?.()
    } catch (err) {
      setReextractProgress(null)
      showToast('error', `Échec de la réextraction : ${err.message || 'Erreur inconnue'}`)
    }
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
  function rowMatchesSearch(row, q) {
    if (!q) return true
    const raw = row.data || {}
    const bailData = row._bailData || {}
    const mods = raw.champs_modifies || {}
    const bailRef = raw.bail_reference || {}
    // Pour un avenant, les champs pertinents peuvent venir : du bail hérité,
    // des champs modifiés par l'avenant, ou de la référence bail_reference
    // renvoyée par l'IA — on cherche dans les trois pour ne rien manquer.
    const searchIn = [
      raw.immeuble, raw.adresse, raw.ville, raw.preneur, raw.bailleur,
      bailData.immeuble, bailData.adresse, bailData.ville, bailData.preneur, bailData.bailleur,
      mods.immeuble, mods.adresse, mods.ville, mods.preneur, mods.bailleur,
      bailRef.immeuble, bailRef.adresse, bailRef.preneur, bailRef.bailleur,
      raw.objet_avenant, raw.sous_location, mods.sous_location,
      row.file_name, row._parentName, row.actif_group,
    ].filter(Boolean).join(' ').toLowerCase()
    return searchIn.includes(q)
  }
  const filtered = q ? displayRows.filter(row => rowMatchesSearch(row, q)) : displayRows

  // Sort top-level bails by actif name or preneur, avenants follow their bail
  const getActifName = row => (row.data?.immeuble || row.data?.adresse || row.file_name || '').toLowerCase()
  const getPreneurName = row => (shortPartyName(row.data?.preneur) || '').toLowerCase()
  const sortedFiltered = (() => {
    const bails = filtered.filter(r => r._level === 0)
    const avMap = {}
    filtered.filter(r => r._level > 0).forEach(r => {
      const pid = r.parent_id || r._parentId
      if (!avMap[pid]) avMap[pid] = []
      avMap[pid].push(r)
    })
    bails.sort((a, b) => {
      const cmp = sortBy === 'preneur'
        ? getPreneurName(a).localeCompare(getPreneurName(b), 'fr')
        : getActifName(a).localeCompare(getActifName(b), 'fr')
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
      {confirmReextract && (
        <ConfirmModal
          title="Réextraire ce document ?"
          message={`Les données actuelles de "${confirmReextract.data?.immeuble || confirmReextract.data?.adresse || confirmReextract.file_name}" seront remplacées par une nouvelle extraction à partir du fichier source déjà attaché. Les avenants éventuels ne sont pas affectés. Cette action écrase toute correction manuelle déjà apportée.`}
          confirmLabel="Réextraire"
          onConfirm={() => handleReextract(confirmReextract)}
          onCancel={() => setConfirmReextract(null)}
        />
      )}
      {confirmAttachReextract && (
        <ConfirmModal
          title="Joindre le fichier et réextraire ?"
          message={`Ce document n'a pas encore de fichier source attaché — probablement une extraction ancienne, faite avec des règles de prompt potentiellement dépassées. Le fichier que tu vas choisir remplacera les données actuelles de "${confirmAttachReextract.data?.immeuble || confirmAttachReextract.data?.adresse || confirmAttachReextract.file_name}" par une extraction avec les règles actuelles. Les avenants éventuels ne sont pas affectés.`}
          confirmLabel="Choisir le fichier"
          onConfirm={() => { openAttachPicker(confirmAttachReextract); setConfirmAttachReextract(null) }}
          onCancel={() => setConfirmAttachReextract(null)}
        />
      )}
      {/* Export errors modal */}
      {exportErrors && exportErrors.length > 0 && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.4)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <div style={{ background: 'var(--surface)', borderRadius: '12px', padding: '28px', maxWidth: '520px', width: '90%', boxShadow: '0 8px 32px rgba(0,0,0,.2)' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '16px' }}>
              <span style={{ fontSize: '20px' }}>⚠️</span>
              <div>
                <div style={{ fontWeight: 700, fontSize: '15px' }}>{exportErrors.length} document{exportErrors.length > 1 ? 's' : ''} non exporté{exportErrors.length > 1 ? 's' : ''}</div>
                <div style={{ fontSize: '12px', color: 'var(--text3)' }}>Ces documents ont été inclus dans l'export avec le statut ⚠</div>
              </div>
            </div>
            <div style={{ maxHeight: '280px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '6px', marginBottom: '20px' }}>
              {exportErrors.map((err, i) => (
                <div key={i} style={{ background: 'var(--danger-bg)', border: '1px solid rgba(176,42,42,.15)', borderRadius: '6px', padding: '8px 12px' }}>
                  <div style={{ fontWeight: 600, fontSize: '13px', color: 'var(--danger)' }}>{err.name}</div>
                  <div style={{ fontSize: '11px', color: 'var(--text2)', marginTop: '2px' }}>{err.reason}</div>
                </div>
              ))}
            </div>
            <button
              onClick={() => setExportErrors(null)}
              style={{ width: '100%', padding: '10px', background: 'var(--accent)', color: '#fff', border: 'none', borderRadius: '8px', fontWeight: 600, cursor: 'pointer', fontSize: '13px' }}>
              Fermer
            </button>
          </div>
        </div>
      )}

      {/* Extraction errors modal */}
      {extractionErrors && extractionErrors.length > 0 && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.5)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <div style={{ background: 'var(--surface)', borderRadius: '12px', padding: '28px', maxWidth: '560px', width: '90%', boxShadow: '0 8px 32px rgba(0,0,0,.25)' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '16px' }}>
              <span style={{ fontSize: '24px' }}>❌</span>
              <div>
                <div style={{ fontWeight: 700, fontSize: '16px' }}>{extractionErrors.length} extraction{extractionErrors.length > 1 ? 's' : ''} en erreur</div>
                <div style={{ fontSize: '12px', color: 'var(--text3)', marginTop: '2px' }}>Ces documents apparaissent dans le dashboard avec un tag "Erreur". Vous pouvez les supprimer et réessayer.</div>
              </div>
            </div>
            <div style={{ maxHeight: '300px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '6px', marginBottom: '20px' }}>
              {extractionErrors.map((err, i) => (
                <div key={i} style={{ background: 'var(--danger-bg)', border: '1px solid rgba(176,42,42,.2)', borderRadius: '8px', padding: '10px 14px' }}>
                  <div style={{ fontWeight: 600, fontSize: '13px', color: 'var(--danger)' }}>📄 {err.name}</div>
                  <div style={{ fontSize: '11px', color: 'var(--text2)', marginTop: '3px', fontFamily: 'monospace' }}>{err.reason}</div>
                </div>
              ))}
            </div>
            <button
              onClick={() => setExtractionErrors(null)}
              style={{ width: '100%', padding: '10px', background: 'var(--danger)', color: '#fff', border: 'none', borderRadius: '8px', fontWeight: 600, cursor: 'pointer', fontSize: '13px' }}>
              Fermer
            </button>
          </div>
        </div>
      )}
      {/* Toolbar */}
      <div className="dash-toolbar">
        <div className="dash-stats">
          {(() => {
            // Réactif à la recherche et au filtre Tous/Baux/Avenants — reflète
            // ce qui est effectivement affiché, pas le total global de la base.
            let bailCount = 0, avenantCount = 0, orphanCount = 0
            tree.forEach(node => {
              if (node.document_type === 'bail') {
                if (filter !== 'avenant' && rowMatchesSearch({ data: node.data, file_name: node.file_name, actif_group: node.actif_group }, q)) bailCount++
                if (filter !== 'bail') {
                  ;(node.avenants || []).forEach(av => {
                    if (rowMatchesSearch({ data: av.data, file_name: av.file_name, actif_group: av.actif_group, _bailData: node.data }, q)) avenantCount++
                  })
                }
              } else if (filter !== 'bail') {
                if (rowMatchesSearch({ data: node.data, file_name: node.file_name, actif_group: node.actif_group }, q)) { avenantCount++; orphanCount++ }
              }
            })
            const isFiltered = !!q || filter !== 'all'
            return (
              <>
                <span className="dash-stat">{bailCount} {bailCount !== 1 ? 'baux' : 'bail'}</span>
                <span className="dash-stat">{avenantCount} avenant{avenantCount !== 1 ? 's' : ''}</span>
                {orphanCount > 0 && (
                  <span className="dash-stat" style={{ color: 'var(--danger)' }} title="Avenants sans bail parent rattaché">
                    dont {orphanCount} orphelin{orphanCount !== 1 ? 's' : ''}
                  </span>
                )}
                {isFiltered && (
                  <span className="dash-stat" style={{ color: 'var(--text3)', fontStyle: 'italic' }}>
                    sur {totalCounts?.bailCount ?? 0} baux / {totalCounts?.avenantCount ?? 0} avenants au total
                  </span>
                )}
              </>
            )
          })()}
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
            <button className="btn" style={{ width: 'auto', padding: '5px 12px', display: 'flex', alignItems: 'center', gap: '5px' }} onClick={() => setShowBulkAttach(true)}>
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/>
              </svg>
              Attacher un dossier
            </button>
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

      {showBulkAttach && (
        <BulkAttachModal
          candidateRows={tree.flatMap(b => [b, ...(b.avenants || [])]).filter(r => !r.storage_path)}
          allRows={tree.flatMap(b => [b, ...(b.avenants || [])])}
          onClose={() => setShowBulkAttach(false)}
          onRefresh={onRefresh}
        />
      )}

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
              onClick={() => { if (sortBy === 'actif') setSortDir(d => d === 'asc' ? 'desc' : 'asc'); else { setSortBy('actif'); setSortDir('asc') } }}>
              Actif / Document
              {sortBy === 'actif' && <span style={{ fontSize: '10px', color: 'var(--text3)' }}>{sortDir === 'asc' ? '↑' : '↓'}</span>}
            </div>
            <div className="dash-th" style={{ gridColumn: '2', cursor: 'pointer', userSelect: 'none', display: 'flex', alignItems: 'center', gap: '5px' }}
              onClick={() => { if (sortBy === 'preneur') setSortDir(d => d === 'asc' ? 'desc' : 'asc'); else { setSortBy('preneur'); setSortDir('asc') } }}>
              Preneur
              {sortBy === 'preneur' && <span style={{ fontSize: '10px', color: 'var(--text3)' }}>{sortDir === 'asc' ? '↑' : '↓'}</span>}
            </div>
            <div className="dash-th" style={{ gridColumn: '3' }}>Type</div>
            <div className="dash-th dash-th-right" style={{ gridColumn: '4' }}>Surface</div>
            <div className="dash-th" style={{ gridColumn: '5' }}>Date effet</div>
            <div className="dash-th" style={{ gridColumn: '6' }}>Date fin</div>
            <div className="dash-th" style={{ gridColumn: '7' }}>Break</div>
            <div className="dash-th dash-th-right" style={{ gridColumn: '8' }}>Loyer HT/HC à la signature</div>
            <div style={{ gridColumn: '9' }}/>
          </div>
          {sortedFiltered.map((row, rowIdx) => {
            // Group header
            if (row._isGroupHeader) return (
              <div key={`grp-${row._groupName}`} style={{
                display: 'grid', gridTemplateColumns: '1fr auto',
                alignItems: 'center', padding: '10px 16px 6px',
                background: 'var(--surface2)', borderBottom: '1px solid var(--border2)',
                marginTop: rowIdx > 0 ? '4px' : 0, position: 'relative',
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span style={{ fontSize: '11px', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '.08em', color: 'var(--accent)' }}>
                    📁 {renamingGroup === row._groupName ? (
                      <input
                        autoFocus
                        defaultValue={row._groupName}
                        onBlur={e => renameGroup(row._groupName, e.target.value)}
                        onKeyDown={e => {
                          if (e.key === 'Enter') renameGroup(row._groupName, e.target.value)
                          if (e.key === 'Escape') setRenamingGroup(null)
                        }}
                        onClick={e => e.stopPropagation()}
                        style={{ fontSize: '11px', fontWeight: 700, letterSpacing: '.08em', color: 'var(--accent)',
                          background: 'transparent', border: 'none', borderBottom: '1px solid var(--accent)',
                          outline: 'none', width: '160px', padding: '0' }}
                      />
                    ) : row._groupName}
                  </span>
                  {renamingGroup !== row._groupName && (
                    <span title="Renommer ce groupe"
                      onClick={e => { e.stopPropagation(); setRenamingGroup(row._groupName) }}
                      style={{ fontSize: '11px', cursor: 'pointer', color: 'var(--text3)', opacity: 0.6 }}>✏️</span>
                  )}
                </div>
                <span style={{ fontSize: '11px', color: 'var(--text3)' }}>{row._groupCount} {row._groupCount > 1 ? 'baux' : 'bail'}</span>
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
                  surface_totale_m2: mods.surface_totale_m2 ?? bailBase.surface_totale_m2,
                  objet_avenant: row.data?.objet_avenant,
                }
              : (row.data || {})
            const isNew = newIds?.includes(row.id)
            const isAv = row.document_type === 'avenant'
            const isOrphan = isAv && !row.parent_id && row._level === 0
            const isExtractionError = d.extraction_error === true
            const breaks = Array.isArray(d.break_options) ? d.break_options : []
            return (
              <div
                key={row.id}
                className={`dash-row${isNew ? ' dash-row-new' : ''}${row._level ? ' dash-row-av' : ''}${isOrphan ? ' dash-row-orphan' : ''}${isExtractionError ? ' dash-row-error' : ''}`}
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
                    {isAv ? (d.objet_avenant || row._parentName || '') : (() => {
                      const v = d.ville || ''
                      const cpMatch = v.match(/(\d{5})/)
                      const cp = cpMatch ? cpMatch[1] : (d.adresse?.match(/(\d{5})/)?.[1] || '')
                      const cityOnly = v.replace(/\d{5}\s*/g, '').replace(/[()]/g, '').trim()
                      const city = cityOnly.toUpperCase()
                      // Add étage if found in surfaces_detail
                      const niveaux = (d.surfaces_detail || [])
                        .filter(r => r.niveau && !( (r.categorie||'').toLowerCase().includes('station') ))
                        .map(r => r.niveau).filter(Boolean)
                      const etage = niveaux.length === 1 ? niveaux[0] : niveaux.length > 1 ? `${niveaux[0]} +${niveaux.length-1}` : ''
                      const loc = cp ? `${city} (${cp})` : city
                      return etage ? `${loc} · ${etage}` : loc
                    })()}
                  </div>
                  {!isAv && (
                    <div style={{ position: 'relative', marginTop: '2px' }}>
                      <span
                        onClick={e => {
                          e.stopPropagation()
                          if (editingActif === row.id) { setEditingActif(null); return }
                          setEditingActifRect(e.currentTarget.getBoundingClientRect())
                          setEditingActif(row.id)
                        }}
                        title="Définir l'actif groupant"
                        style={{ fontSize: '10px', padding: '1px 6px', borderRadius: '4px', cursor: 'pointer',
                          background: row.actif_group ? 'var(--accent-bg)' : 'var(--surface2)',
                          color: row.actif_group ? 'var(--accent)' : 'var(--text3)',
                          border: `1px solid ${row.actif_group ? 'rgba(26,95,168,.2)' : 'var(--border)'}`,
                          fontWeight: row.actif_group ? 600 : 400, display: 'inline-block',
                        }}>
                        {row.actif_group || '+ Actif'}
                      </span>
                      {editingActif === row.id && (
                        <ActifPicker
                          currentValue={row.actif_group || ''}
                          existingGroups={existingGroups}
                          onSave={v => saveActifGroup(row.id, v)}
                          onClose={() => setEditingActif(null)}
                          anchorRect={editingActifRect}
                        />
                      )}
                    </div>
                  )}
                </div>

                {/* Preneur */}
                <div className="dash-td" style={{ alignItems: 'flex-start', paddingTop: '13px' }}>
                  <span style={{ fontSize: '12px', color: 'var(--text)', fontWeight: 500, whiteSpace: 'normal', wordBreak: 'break-word', lineHeight: 1.35, display: 'block' }}>
                    {shortPartyName(d.preneur)?.toUpperCase() || '—'}
                  </span>
                </div>

                {/* Type */}
                <div className="dash-td" style={{ alignItems: 'flex-start', paddingTop: '13px', flexDirection: 'column', gap: '3px' }}>
                  <span className={`dash-tag ${isAv ? 'dash-tag-av' : 'dash-tag-bail'}`}>
                    {isAv ? 'Avenant' : 'Bail'}
                  </span>
                  {isExtractionError && (
                    <span title={d.error_message || 'Erreur lors de l\'extraction'} style={{ fontSize: '10px', background: 'var(--danger-bg)', color: 'var(--danger)', border: '1px solid rgba(176,42,42,.2)', borderRadius: '4px', padding: '1px 5px', fontWeight: 600, marginTop: '2px', display: 'block', cursor: 'help' }}>
                      ❌ Erreur extraction
                    </span>
                  )}
                  {isOrphan && !isExtractionError && (
                    <span title="Bail parent manquant ou en erreur" style={{ fontSize: '10px', background: 'var(--danger-bg)', color: 'var(--danger)', border: '1px solid rgba(176,42,42,.2)', borderRadius: '4px', padding: '1px 5px', fontWeight: 600, marginTop: '2px', display: 'block' }}>
                      ⚠ Bail manquant
                    </span>
                  )}
                </div>

                {/* Surface */}
                <div className="dash-td dash-td-right" style={{ alignItems: 'flex-start', paddingTop: '13px' }}>
                  <span style={{ fontSize: '12px', color: 'var(--text2)', lineHeight: 1.4 }}>
                    {d.surface_totale_m2 ? `${d.surface_totale_m2} m²` : '—'}
                  </span>
                </div>

                {/* Date effet */}
                <div className="dash-td" style={{ alignItems: 'flex-start', paddingTop: '13px' }}>
                  <span style={{ fontSize: '12px', color: 'var(--text2)', lineHeight: 1.4 }}>{normalizeDate(d.date_effet) || '—'}</span>
                </div>

                {/* Date fin */}
                <div className="dash-td" style={{ alignItems: 'flex-start', paddingTop: '13px' }}>
                  <span style={{ fontSize: '12px', color: 'var(--text2)', lineHeight: 1.4 }}>{normalizeDate(d.date_fin) || '—'}</span>
                </div>

                {/* Break */}
                <div className="dash-td" style={{ alignItems: 'flex-start', paddingTop: '13px' }}>
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
                <div className="dash-td dash-td-right" style={{ alignItems: 'flex-start', paddingTop: '13px' }}>
                  {d.loyer_signature_montant ? (
                    <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text)', letterSpacing: '-0.01em' }}>
                      {fmtEur(d.loyer_signature_montant)}
                    </span>
                  ) : <span style={{ fontSize: '12px', color: 'var(--text3)' }}>—</span>}
                </div>

                {/* Actions */}
                <div className="dash-td dash-td-actions" onClick={e => e.stopPropagation()}>
                  {!isAv && avenantUpload[row.id] && (
                    <span
                      title={avenantUpload[row.id].state === 'error' ? avenantUpload[row.id].error : ''}
                      style={{
                        fontSize: '10px', fontWeight: 600, padding: '2px 6px', borderRadius: '4px', whiteSpace: 'nowrap',
                        background: avenantUpload[row.id].state === 'error' ? 'var(--danger-bg)' : 'var(--accent-bg)',
                        color: avenantUpload[row.id].state === 'error' ? 'var(--danger)' : 'var(--accent)',
                        cursor: avenantUpload[row.id].state === 'error' ? 'help' : 'default',
                      }}>
                      {avenantUpload[row.id].state === 'compressing'
                        ? `Compression${avenantUpload[row.id].total ? ` ${avenantUpload[row.id].current}/${avenantUpload[row.id].total}` : '…'}`
                        : avenantUpload[row.id].state === 'loading' ? 'Traitement…'
                        : '❌ Erreur'}
                    </span>
                  )}
                  {isAv && avenantUpload[row.id] && (
                    <span
                      title={avenantUpload[row.id].state === 'error' ? avenantUpload[row.id].error : ''}
                      style={{
                        fontSize: '10px', fontWeight: 600, padding: '2px 6px', borderRadius: '4px', whiteSpace: 'nowrap',
                        background: avenantUpload[row.id].state === 'error' ? 'var(--danger-bg)' : 'var(--accent-bg)',
                        color: avenantUpload[row.id].state === 'error' ? 'var(--danger)' : 'var(--accent)',
                        cursor: avenantUpload[row.id].state === 'error' ? 'help' : 'default',
                      }}>
                      {avenantUpload[row.id].state === 'compressing'
                        ? `Compression${avenantUpload[row.id].total ? ` ${avenantUpload[row.id].current}/${avenantUpload[row.id].total}` : '…'}`
                        : avenantUpload[row.id].state === 'loading' ? 'Traitement…'
                        : '❌ Erreur'}
                    </span>
                  )}
                  {!isAv && (
                    <button className="dash-action-btn" style={{ opacity: 1 }} onClick={e => { e.stopPropagation(); openAvenantPicker(row) }} title="Ajouter un avenant à ce bail">
                      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                    </button>
                  )}
                  {row.storage_path ? (
                    <button className="dash-action-btn" style={{ opacity: 1 }} onClick={e => { e.stopPropagation(); viewSourceFile(row) }} title="Voir le fichier source">
                      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
                    </button>
                  ) : (
                    <button className="dash-action-btn" style={{ opacity: 1 }} onClick={e => { e.stopPropagation(); setConfirmAttachReextract(row) }} title="Joindre le fichier source et réextraire (remplace les données avec le prompt actuel)">
                      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/></svg>
                    </button>
                  )}
                  {row.storage_path && (
                    <button className="dash-action-btn" style={{ opacity: 1 }} onClick={e => { e.stopPropagation(); setConfirmReextract(row) }} title="Réextraire (remplace les données à partir du fichier source, sans toucher aux avenants)">
                      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 12c0 4.97-4.03 9-9 9s-9-4.03-9-9 4.03-9 9-9c1.5 0 2.91.37 4.15 1.02" /><polyline points="17 3 21 3 21 7"/><path d="M21 3l-8.15 8.15"/></svg>
                    </button>
                  )}
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
      <input
        ref={avenantInputRef}
        type="file"
        accept=".pdf,.docx"
        multiple
        style={{ display: 'none' }}
        onChange={handleAvenantFile}
      />
      <input
        ref={attachInputRef}
        type="file"
        accept=".pdf,.docx"
        style={{ display: 'none' }}
        onChange={handleAttachFile}
      />
      {avenantBatchProgress && (
        <div style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,.5)', zIndex: 5000,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }}>
          <div style={{
            background: 'var(--surface)', borderRadius: '14px', padding: '32px 40px',
            minWidth: '360px', maxWidth: '440px', textAlign: 'center', boxShadow: '0 16px 48px rgba(0,0,0,.35)',
          }}>
            <div style={{
              width: '36px', height: '36px', margin: '0 auto 16px', borderRadius: '50%',
              border: '3px solid var(--border2)', borderTopColor: 'var(--accent)',
              animation: 'spin 0.8s linear infinite',
            }} />
            <div style={{ fontWeight: 700, fontSize: '15px', marginBottom: '6px' }}>
              Ajout d'avenant{avenantBatchProgress.total > 1 ? 's' : ''} à « {avenantBatchProgress.bailLabel} »
            </div>
            <div style={{ fontSize: '13px', color: 'var(--text2)', marginBottom: '10px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              Fichier {avenantBatchProgress.current}/{avenantBatchProgress.total} : {avenantBatchProgress.fileName}
            </div>
            <div style={{ fontSize: '12px', color: 'var(--text3)', marginBottom: '14px' }}>
              {avenantBatchProgress.state === 'compressing'
                ? `Compression${avenantBatchProgress.progTotal ? ` (page ${avenantBatchProgress.progCurrent}/${avenantBatchProgress.progTotal})` : '…'}`
                : 'Extraction en cours…'}
            </div>
            <div className="progress-track" style={{ margin: 0 }}><div className="progress-bar active" /></div>
          </div>
        </div>
      )}
      {reextractProgress && (
        <div style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,.5)', zIndex: 5000,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }}>
          <div style={{
            background: 'var(--surface)', borderRadius: '14px', padding: '32px 40px',
            minWidth: '360px', maxWidth: '440px', textAlign: 'center', boxShadow: '0 16px 48px rgba(0,0,0,.35)',
          }}>
            <div style={{
              width: '36px', height: '36px', margin: '0 auto 16px', borderRadius: '50%',
              border: '3px solid var(--border2)', borderTopColor: 'var(--accent)',
              animation: 'spin 0.8s linear infinite',
            }} />
            <div style={{ fontWeight: 700, fontSize: '15px', marginBottom: '6px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              Réextraction de « {reextractProgress.label} »
            </div>
            <div style={{ fontSize: '12px', color: 'var(--text3)', marginBottom: '14px' }}>
              {reextractProgress.state === 'downloading' && 'Téléchargement du fichier source…'}
              {reextractProgress.state === 'compressing' && `Compression${reextractProgress.progTotal ? ` (page ${reextractProgress.progCurrent}/${reextractProgress.progTotal})` : '…'}`}
              {reextractProgress.state === 'loading' && 'Extraction en cours…'}
            </div>
            <div className="progress-track" style={{ margin: 0 }}><div className="progress-bar active" /></div>
          </div>
        </div>
      )}
      {toast && (
        <div
          onClick={() => setToast(null)}
          style={{
            position: 'fixed', bottom: '24px', right: '24px', zIndex: 2000,
            display: 'flex', alignItems: 'center', gap: '10px',
            padding: '12px 16px', borderRadius: '8px', cursor: 'pointer',
            background: toast.type === 'error' ? 'var(--danger-bg)' : 'var(--surface)',
            border: `1px solid ${toast.type === 'error' ? 'rgba(176,42,42,.25)' : 'var(--border2)'}`,
            boxShadow: '0 8px 24px rgba(0,0,0,.15)', maxWidth: '360px',
            animation: 'toastIn .25s ease-out',
          }}>
          <span style={{ fontSize: '16px', flexShrink: 0 }}>{toast.type === 'error' ? '❌' : '✅'}</span>
          <span style={{ fontSize: '13px', color: toast.type === 'error' ? 'var(--danger)' : 'var(--text)', lineHeight: 1.4 }}>
            {toast.message}
          </span>
        </div>
      )}
    </div>
  )
}

// ─── App ─────────────────────────────────────────────────────────────────────

export default function App() {
  const [files,        setFiles]        = useState([])
  const dirActifGroupsRef = useRef({})
  const [statuses,     setStatuses]     = useState([])
  const [loading,      setLoading]      = useState(false)
  const [activeItem,   setActiveItem]   = useState(null)
  const [history,      setHistory]      = useState([])
  const [histLoaded,   setHistLoaded]   = useState(false)
  const [totalCounts,  setTotalCounts]  = useState({ bailCount: 0, avenantCount: 0, orphanCount: 0 })
  const [tab,          setTab]          = useState('history')
  const [showAddModal, setShowAddModal] = useState(false)
  const [showEtatLocatifMenu, setShowEtatLocatifMenu] = useState(false)
  const [etatLocatifBuilding, setEtatLocatifBuilding] = useState(null)
  const [showQualityCheck, setShowQualityCheck] = useState(false)
  // Recherche/filtre du dashboard remontés ici (plutôt que locaux à Dashboard)
  // pour survivre à la navigation vers une fiche détail et retour.
  const [dashSearch, setDashSearch] = useState('')
  const [dashFilter, setDashFilter] = useState('all')
  const [docTypes,     setDocTypes]     = useState([])     // 'bail'|'avenant'|'' per file
  const [fileOrder,    setFileOrder]    = useState([])     // indices ordonnés
  const [detecting,    setDetecting]    = useState(false)  // détection en cours
  const [avenantLinks, setAvenantLinks] = useState({})     // index -> parentId
  const [pertinents,   setPertinents]   = useState([])     // bool per file
  const [raisons,      setRaisons]      = useState([])     // raison non pertinent
  const [lastError,    setLastError]    = useState('')
  const [newIds,       setNewIds]       = useState([])   // ids extraits dans le batch courant
  const [compressing,  setCompressing]  = useState(null) // { name, current, total } | null

  function buildTree(rows) {
    const bails    = rows.filter(r => r.document_type === 'bail')
    const avenants = rows.filter(r => r.document_type === 'avenant' && r.parent_id)
    const orphans  = rows.filter(r => r.document_type === 'avenant' && !r.parent_id)
    return [...bails.map(b => ({ ...b, avenants: avenants.filter(a => a.parent_id === b.id) })), ...orphans]
  }

  // Le rendu du tableau charge jusqu'à RENDER_LIMIT lignes (largement au-dessus
  // du volume actuel). Le COMPTEUR affiché, lui, vient d'une requête de comptage
  // exact séparée (fetchTotalCounts), donc il reste juste même si RENDER_LIMIT
  // était un jour dépassé.
  const RENDER_LIMIT = 2000

  async function fetchAllHistory() {
    const { data: rows } = await supabase.from('extractions')
      .select('id, file_name, created_at, data, document_type, parent_id, actif_group, storage_path')
      .order('created_at', { ascending: false }).limit(RENDER_LIMIT)
    return rows ? buildTree(rows) : []
  }

  async function fetchTotalCounts() {
    const [bailRes, avenantRes, orphanRes] = await Promise.all([
      supabase.from('extractions').select('*', { count: 'exact', head: true }).eq('document_type', 'bail'),
      supabase.from('extractions').select('*', { count: 'exact', head: true }).eq('document_type', 'avenant'),
      supabase.from('extractions').select('*', { count: 'exact', head: true }).eq('document_type', 'avenant').is('parent_id', null),
    ])
    return {
      bailCount: bailRes.count || 0,
      avenantCount: avenantRes.count || 0,
      orphanCount: orphanRes.count || 0,
    }
  }

  async function loadHistory() {
    if (histLoaded) return
    const [tree, counts] = await Promise.all([fetchAllHistory(), fetchTotalCounts()])
    setHistory(tree)
    setTotalCounts(counts)
    setHistLoaded(true)
  }

  // ─── Navigation par URL (API History native — pas de librairie de routage) ──
  // /               → dashboard
  // /bail/{id}      → fiche détail d'un bail ou avenant
  // /etat-locatif/{immeuble encodé} → état locatif d'un actif groupant
  function navigate(path) {
    if (window.location.pathname !== path) window.history.pushState({}, '', path)
  }

  async function loadItemById(id) {
    const { data } = await supabase.from('extractions')
      .select('id, file_name, created_at, data, document_type, parent_id, actif_group, storage_path')
      .eq('id', id).single()
    return data || null
  }

  async function applyUrlState() {
    const path = window.location.pathname
    const bailMatch = path.match(/^\/bail\/([^/]+)/)
    const etatMatch = path.match(/^\/etat-locatif\/([^/]+)/)
    if (bailMatch) {
      const item = await loadItemById(decodeURIComponent(bailMatch[1]))
      setEtatLocatifBuilding(null)
      setActiveItem(item) // null si id introuvable/supprimé → retombe proprement sur le dashboard
    } else if (etatMatch) {
      setActiveItem(null)
      setEtatLocatifBuilding(decodeURIComponent(etatMatch[1]))
    } else {
      setActiveItem(null)
      setEtatLocatifBuilding(null)
    }
  }

  // Lien direct/partagé ouvert à froid : appliquer l'URL au montage
  useEffect(() => { applyUrlState() }, [])

  // Boutons précédent/suivant du navigateur
  useEffect(() => {
    window.addEventListener('popstate', applyUrlState)
    return () => window.removeEventListener('popstate', applyUrlState)
  }, [])

  // Le Dashboard est désormais la seule page (plus d'onglet "Extraire" séparé
  // à cliquer en premier) — il faut donc charger l'historique dès le montage.
  useEffect(() => { loadHistory() }, [])

  // Clic sur un lot dans l'état locatif → ouvre le détail du bail, ferme la modale
  useEffect(() => {
    function handler(e) {
      setEtatLocatifBuilding(null)
      setActiveItem(e.detail)
      navigate(`/bail/${e.detail.id}`)
    }
    window.addEventListener('etatlocatif-select', handler)
    return () => window.removeEventListener('etatlocatif-select', handler)
  }, [])

  // Fermeture du menu "État locatif" au clic extérieur
  useEffect(() => {
    if (!showEtatLocatifMenu) return
    const handler = () => setShowEtatLocatifMenu(false)
    document.addEventListener('click', handler)
    return () => document.removeEventListener('click', handler)
  }, [showEtatLocatifMenu])

  const buildingGroups = useMemo(() => {
    const map = {}
    history.forEach(row => {
      if (row.document_type !== 'bail' || !row.actif_group) return
      map[row.actif_group] = (map[row.actif_group] || 0) + 1
    })
    return Object.entries(map).map(([name, count]) => ({ name, count })).sort((a, b) => a.name.localeCompare(b.name))
  }, [history])

  // Rafraîchissement forcé (ignore le cache histLoaded) — utilisé après un ajout
  // ponctuel depuis le dashboard (ex. bouton "+ Avenant"), où loadHistory() seul
  // ne rechargerait rien puisque histLoaded est déjà à true.
  async function refreshHistoryNow() {
    const [tree, counts] = await Promise.all([fetchAllHistory(), fetchTotalCounts()])
    setHistory(tree)
    setTotalCounts(counts)
    setHistLoaded(true)
    // Si la fiche détail actuellement affichée correspond à un document rafraîchi,
    // la resynchroniser aussi — sinon elle reste figée sur l'ancienne version tant
    // qu'on ne revient pas au dashboard (ex: après 🔄 Réextraire ou 📎 Attacher).
    setActiveItem(prev => {
      if (!prev) return prev
      const flat = tree.flatMap(b => [b, ...(b.avenants || [])])
      const fresh = flat.find(r => r.id === prev.id)
      return fresh || prev
    })
  }

  async function switchTab(t) {
    setTab(t)
    if (t === 'history') {
      // Forcer rechargement depuis Supabase directement
      const [tree, counts] = await Promise.all([fetchAllHistory(), fetchTotalCounts()])
      setHistory(tree)
      setTotalCounts(counts)
      setHistLoaded(true)
    }
  }
  function setStatus(i, state, error) { setStatuses(prev => { const n = [...prev]; n[i] = { state, error }; return n }) }

  async function saveExtraction(file, extracted, docType, parentId, actifGroup = null) {
    const { data: saved } = await supabase.from('extractions')
      .insert({ file_name: file.name, data: extracted, document_type: docType, parent_id: parentId || null, actif_group: actifGroup || null })
      .select().single()
    if (saved?.id) await uploadSourceFile(saved.id, file) // on attend la fin pour éviter un rafraîchissement prématuré du dashboard
    return saved
  }

  // Détection automatique déclenchée au drop
  async function detectFiles(newFiles, offset = 0, dirAutoLinks = {}, dirActifGroups = {}) {
    setDetecting(true)
    // Store dir-based actif groups (absolute indices)
    Object.entries(dirActifGroups).forEach(([k, v]) => { dirActifGroupsRef.current[parseInt(k) + offset] = v })
    const types      = new Array(newFiles.length).fill('')
    const pertinents = new Array(newFiles.length).fill(null)
    const raisons    = new Array(newFiles.length).fill('')
    const detectData = new Array(newFiles.length).fill(null)
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
        // Update global state at offset position
        setDocTypes(prev => { const n = [...prev]; n[offset + i] = types[i]; return n })
        setPertinents(prev => { const n = [...prev]; n[offset + i] = pertinents[i]; return n })
        setRaisons(prev => { const n = [...prev]; n[offset + i] = raisons[i]; return n })
      }))
    }
    const bailIdx    = types.map((t,i) => t === 'bail'    ? i : -1).filter(i => i >= 0).map(i => i + offset)
    const avenantIdx = types.map((t,i) => t === 'avenant' ? i : -1).filter(i => i >= 0).map(i => i + offset)
    setFileOrder(prev => {
      const bails = prev.filter(i => i < offset || types[i - offset] === 'bail')
      const avs   = prev.filter(i => i >= offset && types[i - offset] === 'avenant')
      return [...bails, ...avs]
    })
    const existingBails = history.filter(h => h.document_type === 'bail')
    const batchBails = bailIdx
      .filter(i => pertinents[i - offset] !== false)
      .map(i => ({
        id: `batch-${i}`,
        file_name: newFiles[i - offset].name,
        data: { preneur: detectData[i - offset]?.preneur, bailleur: detectData[i - offset]?.bailleur, adresse: detectData[i - offset]?.adresse, immeuble: detectData[i - offset]?.immeuble }
      }))
    const allBailsForMatch = [...existingBails, ...batchBails]
    const autoLinks = {}

    // First: apply directory-based links (most reliable)
    Object.entries(dirAutoLinks).forEach(([avRelIdx, bailRelIdx]) => {
      const avAbsIdx = offset + parseInt(avRelIdx)
      // Find the batch bail that corresponds to bailRelIdx
      const bailFile = newFiles[parseInt(bailRelIdx)]
      const bailBatch = batchBails.find(b => b.file_name === bailFile?.name)
      if (bailBatch) autoLinks[avAbsIdx] = bailBatch.id
    })

    // Then: AI-based matching for avenants not already linked by directory
    avenantIdx
      .filter(i => pertinents[i - offset] !== false && !autoLinks[i])
      .forEach(i => {
        if (allBailsForMatch.length === 1) {
          autoLinks[i] = allBailsForMatch[0].id
        } else if (allBailsForMatch.length > 1 && detectData[i - offset]) {
          const match = findBestMatch(detectData[i - offset], allBailsForMatch)
          if (match) autoLinks[i] = match.item.id
        }
      })
    setAvenantLinks(prev => ({ ...prev, ...autoLinks }))
    setDetecting(false)
  }

  async function handleFiles(newFiles, dirAutoLinks = {}, dirActifGroups = {}) {
    let arr = Array.from(newFiles)

    // Compression préventive des PDF volumineux (scans / "Print to PDF")
    // avant toute détection ou extraction — évite les erreurs de taille en aval.
    const heavy = arr.filter(f => f.name.toLowerCase().endsWith('.pdf') && f.size > PDF_COMPRESS_THRESHOLD)
    if (heavy.length > 0) {
      for (const f of heavy) {
        const idx = arr.indexOf(f)
        setCompressing({ name: f.name, current: 0, total: 0 })
        const compressed = await compressPdfIfNeeded(f, (current, total) => {
          setCompressing({ name: f.name, current, total })
        })
        arr[idx] = compressed
      }
      setCompressing(null)
    }

    setFiles(prev => {
      const combined = [...prev, ...arr]
      const offset = prev.length
      setDocTypes(pt => [...pt, ...arr.map(() => '')])
      setFileOrder(po => [...po, ...arr.map((_, i) => offset + i)])
      setStatuses(ps => [...ps, ...arr.map(() => ({}))])
      setPertinents(pp => [...pp, ...arr.map(() => null)])
      setRaisons(pr => [...pr, ...arr.map(() => '')])
      setLastError('')
      // Apply directory-based auto-links (relative indices → absolute)
      if (Object.keys(dirAutoLinks).length > 0) {
        const absLinks = {}
        Object.entries(dirAutoLinks).forEach(([avRelIdx, bailRelIdx]) => {
          absLinks[offset + parseInt(avRelIdx)] = `dir-${offset + parseInt(bailRelIdx)}`
        })
        setAvenantLinks(prev => ({ ...prev, ...absLinks }))
      }
      detectFiles(arr, offset, dirAutoLinks, dirActifGroups)
      return combined
    })
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

  // Concurrency queue: run tasks with max N concurrent
  async function runWithConcurrency(items, maxConcurrent, taskFn) {
    const queue = [...items]
    const workers = Array(Math.min(maxConcurrent, items.length)).fill(null).map(async () => {
      while (queue.length > 0) {
        const item = queue.shift()
        if (item !== undefined) await taskFn(item)
      }
    })
    await Promise.all(workers)
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

    // 1. Extraire les baux d'abord (max 4 en parallèle)
    await runWithConcurrency(bailIndices, 4, async (i) => {
      try {
        setStatus(i, 'loading')
        if (files[i].size > 30 * 1024 * 1024) throw new Error(`Fichier trop volumineux (${Math.round(files[i].size/1024/1024)} Mo > 30 Mo) — compressez le PDF avant de déposer.`)
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
        const saved = await saveExtraction(files[i], extracted, 'bail', null, dirActifGroupsRef.current[i] || null)
        if (saved) {
          const bwa = { ...saved, avenants: [] }
          availableBails.push(bwa)
          setNewIds(prev => [...prev, saved.id])
          // Ne pas setHistory ici — on recharge tout à la fin
        }
        setStatus(i, 'done')
      } catch (e) {
        setStatus(i, 'error', e.message); setLastError(e.message)
        extractionErrorsList.push({ name: files[i]?.name || `Fichier ${i+1}`, reason: e.message || 'Erreur inconnue' })
        try {
          await supabase.from('extractions').insert({ file_name: files[i]?.name, data: { extraction_error: true, error_message: e.message }, document_type: 'bail', parent_id: null, actif_group: dirActifGroupsRef.current[i] || null })
        } catch (_) {}
      }
    })

    // 2. Extraire les avenants et sauvegarder directement avec le bail lié choisi
    let lastSaved = null
    const extractionErrorsList = [] // accumulate errors during extraction
    await runWithConcurrency(avenantIndices, 3, async (i) => {
      try {
        setStatus(i, 'loading')
        if (files[i].size > 30 * 1024 * 1024) throw new Error(`Fichier trop volumineux (${Math.round(files[i].size/1024/1024)} Mo > 30 Mo) — compressez le PDF avant de déposer.`)
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
        // Résoudre batch- et dir- id en vrai id
        let parentId = avenantLinks[i] || null
        if (parentId && parentId.startsWith('batch-')) {
          const batchIdx = parseInt(parentId.replace('batch-', ''))
          const realBail = availableBails.find(b => b.file_name === files[batchIdx]?.name)
          parentId = realBail?.id || null
        } else if (parentId && parentId.startsWith('dir-')) {
          const dirIdx = parseInt(parentId.replace('dir-', ''))
          const realBail = availableBails.find(b => b.file_name === files[dirIdx]?.name)
          parentId = realBail?.id || null
        }
        const saved = await saveExtraction(files[i], extracted, 'avenant', parentId, dirActifGroupsRef.current[i] || null)
        if (saved) {
          lastSaved = saved
          setNewIds(prev => [...prev, saved.id])
          // Ne pas setHistory ici — on recharge tout à la fin
        }
        setStatus(i, 'done')
      } catch (e) {
        setStatus(i, 'error', e.message); setLastError(e.message)
        extractionErrorsList.push({ name: files[i]?.name || `Fichier ${i+1}`, reason: e.message || 'Erreur inconnue' })
        try {
          await supabase.from('extractions').insert({ file_name: files[i]?.name, data: { extraction_error: true, error_message: e.message }, document_type: 'avenant', parent_id: null, actif_group: dirActifGroupsRef.current[i] || null })
        } catch (_) {}
      }
    })

    setLoading(false)
    // Recharger l'historique complet depuis Supabase
    setHistLoaded(false)
    const [freshTree, freshCounts] = await Promise.all([fetchAllHistory(), fetchTotalCounts()])
    setHistory(freshTree)
    setTotalCounts(freshCounts)
    setHistLoaded(true)
    setTab('history')
    setShowAddModal(false)
    handleClear()
    // Modale erreurs d'extraction
    if (extractionErrorsList.length > 0) setExtractionErrors(extractionErrorsList)
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

  const contentRef = useRef()
  useEffect(() => {
    requestAnimationFrame(() => {
      if (contentRef.current) contentRef.current.scrollTop = 0
      window.scrollTo(0, 0)
    })
  }, [activeItem?.id])
  const resultTitle = d.immeuble || d.adresse || activeItem?.file_name || ''
  const shortName = s => s?.split(',')[0]?.split('(')[0]?.split(' SAS')[0]?.split(' SA ')[0]?.trim()
  const resultSub = [shortName(d.preneur), shortName(d.bailleur), d.date_signature ? `Signé le ${d.date_signature}` : null].filter(Boolean).join(' · ')

  return (
    <>
      <div className="app">
        <header className="topbar">
          <div onClick={() => { setActiveItem(null); navigate('/'); switchTab('history') }} style={{ display: 'flex', alignItems: 'center', gap: '9px', cursor: 'pointer' }}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14 2 14 8 20 8"/>
              <line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>
            </svg>
            Lease Reader
          </div>

          <div style={{ position: 'relative', marginLeft: '24px' }} onClick={e => e.stopPropagation()}>
            <button
              onClick={() => setShowEtatLocatifMenu(v => !v)}
              style={{
                display: 'flex', alignItems: 'center', gap: '7px', background: 'rgba(255,255,255,0.08)',
                border: '1px solid rgba(255,255,255,0.15)', color: '#fff', fontSize: '13px', fontWeight: 600,
                padding: '6px 12px', borderRadius: '6px', cursor: 'pointer',
              }}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <rect x="3" y="3" width="18" height="18" rx="1"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="12" y1="3" x2="12" y2="21"/>
              </svg>
              État locatif
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="6 9 12 15 18 9"/></svg>
            </button>
            {showEtatLocatifMenu && (
              <div style={{
                position: 'absolute', top: '100%', left: 0, marginTop: '6px', background: 'var(--surface)',
                border: '1px solid var(--border2)', borderRadius: '8px', boxShadow: '0 8px 24px rgba(0,0,0,.25)',
                width: '260px', maxHeight: '320px', overflowY: 'auto', zIndex: 500,
              }}>
                {buildingGroups.length === 0 ? (
                  <div style={{ padding: '14px 12px', fontSize: '12px', color: 'var(--text3)', fontStyle: 'italic' }}>
                    Aucun actif groupant défini pour l'instant
                  </div>
                ) : buildingGroups.map(g => (
                  <div
                    key={g.name}
                    onClick={() => { setEtatLocatifBuilding(g.name); setShowEtatLocatifMenu(false); navigate(`/etat-locatif/${encodeURIComponent(g.name)}`) }}
                    style={{ padding: '9px 12px', fontSize: '13px', cursor: 'pointer', borderBottom: '1px solid var(--border)', color: 'var(--text)' }}
                    onMouseEnter={e => e.currentTarget.style.background = 'var(--accent-bg)'}
                    onMouseLeave={e => e.currentTarget.style.background = ''}>
                    {g.name} <span style={{ color: 'var(--text3)', fontSize: '11px' }}>({g.count} {g.count === 1 ? 'bail' : 'baux'})</span>
                  </div>
                ))}
              </div>
            )}
          </div>

          <button
            onClick={() => setShowQualityCheck(true)}
            style={{
              display: 'flex', alignItems: 'center', gap: '7px', background: 'rgba(255,255,255,0.08)',
              border: '1px solid rgba(255,255,255,0.15)', color: '#fff', fontSize: '13px', fontWeight: 600,
              padding: '6px 12px', borderRadius: '6px', cursor: 'pointer', marginLeft: '10px',
            }}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M9 12l2 2 4-4"/><path d="M21 12c0 4.97-4.03 9-9 9s-9-4.03-9-9 4.03-9 9-9c1.5 0 2.91.37 4.15 1.02"/><path d="M22 4L12 14.01l-3-3"/>
            </svg>
            Contrôle qualité
          </button>
        </header>

        {showQualityCheck && (
          <QualityCheckModal
            bails={history.filter(row => row.document_type === 'bail')}
            onClose={() => setShowQualityCheck(false)}
            onSelect={row => { setShowQualityCheck(false); setActiveItem(row); navigate(`/bail/${row.id}`) }}
            onDismiss={async (rowId, dismissed) => {
              const row = history.find(b => b.id === rowId)
              if (!row) return
              const newData = { ...row.data, _qc_dismissed: dismissed }
              await supabase.from('extractions').update({ data: newData }).eq('id', rowId)
              setHistory(prev => prev.map(b => b.id === rowId ? { ...b, data: newData } : b))
              if (activeItem?.id === rowId) setActiveItem(prev => ({ ...prev, data: newData }))
            }}
            onFixAnniversary={async (rowId, dates) => {
              const row = history.find(b => b.id === rowId)
              if (!row) return
              const toFix = new Set(dates)
              const newBreaks = (row.data?.break_options || []).map(b => {
                if (typeof b === 'string' && toFix.has(b.trim())) {
                  const d = parseFR(b.trim())
                  if (!d) return b
                  d.setDate(d.getDate() - 1)
                  return fmtFR(d)
                }
                return b
              })
              const newData = { ...row.data, break_options: newBreaks }
              await supabase.from('extractions').update({ data: newData }).eq('id', rowId)
              setHistory(prev => prev.map(b => b.id === rowId ? { ...b, data: newData } : b))
              if (activeItem?.id === rowId) setActiveItem(prev => ({ ...prev, data: newData }))
            }}
          />
        )}

        {etatLocatifBuilding && (
          <EtatLocatifModal
            building={etatLocatifBuilding}
            bails={history.filter(row => row.document_type === 'bail' && row.actif_group === etatLocatifBuilding)}
            onClose={() => { setEtatLocatifBuilding(null); navigate('/') }}
          />
        )}

        <main className="main">
          {activeItem && (() => {
            // Navigation bail ↔ avenants : calculée ici plutôt que dans
            // ResultsView, qui n'a pas accès à `history` (la liste complète).
            const toSortable = s => {
              const m = String(s || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/)
              return m ? `${m[3]}-${m[2]}-${m[1]}` : String(s || '')
            }
            const avKey = av => toSortable(av.data?.date_effet_avenant || av.data?.date_signature_avenant || av.created_at)
            let avNav = null
            if (activeItem.document_type === 'bail' && activeItem.avenants?.length > 0) {
              const sorted = [...activeItem.avenants].sort((a, b) => avKey(a).localeCompare(avKey(b)))
              avNav = { next: sorted[0], nextLabel: sorted.length > 1 ? `Avenant 1/${sorted.length} →` : 'Avenant →' }
            } else if (activeItem.document_type === 'avenant' && activeItem.parent_id) {
              const parentBail = history.find(b => b.id === activeItem.parent_id)
              if (parentBail) {
                const sorted = [...(parentBail.avenants || [])].sort((a, b) => avKey(a).localeCompare(avKey(b)))
                const idx = sorted.findIndex(a => a.id === activeItem.id)
                avNav = {
                  prev: idx > 0 ? sorted[idx - 1] : parentBail,
                  prevLabel: idx > 0 ? `← Avenant ${idx}/${sorted.length}` : '← Bail',
                  next: idx >= 0 && idx < sorted.length - 1 ? sorted[idx + 1] : null,
                  nextLabel: idx >= 0 && idx < sorted.length - 1 ? `Avenant ${idx + 2}/${sorted.length} →` : null,
                }
              }
            }
            return (
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
                <button className="btn back" onClick={() => { setActiveItem(null); navigate('/') }}>← Retour au dashboard</button>
                {avNav?.prev && (
                  <button className="btn" onClick={() => { setActiveItem(avNav.prev); navigate(`/bail/${avNav.prev.id}`) }} title="Document précédent">
                    {avNav.prevLabel}
                  </button>
                )}
                {avNav?.next && (
                  <button className="btn" onClick={() => { setActiveItem(avNav.next); navigate(`/bail/${avNav.next.id}`) }} title="Document suivant">
                    {avNav.nextLabel}
                  </button>
                )}
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
            )
          })()}

          <div className="content" ref={contentRef}>
            {activeItem ? (
              <ResultsView item={activeItem} />
            ) : (
              <>
                <Dashboard
                  tree={history}
                  totalCounts={totalCounts}
                  onSelect={item => { setActiveItem(item); navigate(`/bail/${item.id}`) }}
                  onDelete={handleDeleteItem}
                  onClear={handleClearHistory}
                  onExportAll={() => exportAllToExcel(history, setExportErrors)}
                  newIds={newIds}
                  onRefresh={refreshHistoryNow}
                  onNewAvenant={id => setNewIds(prev => [...prev, id])}
                  filter={dashFilter}
                  setFilter={setDashFilter}
                  search={dashSearch}
                  setSearch={setDashSearch}
                  onUpdateActif={(id, value) => {
                    setHistory(prev => prev.map(b => {
                      if (b.id === id) return { ...b, actif_group: value || null }
                      return { ...b, avenants: (b.avenants || []).map(a => a.id === id ? { ...a, actif_group: value || null } : a) }
                    }))
                  }}
                />

                {/* Bouton flottant d'ajout */}
                {!showAddModal && (
                  <button
                    onClick={() => setShowAddModal(true)}
                    title="Ajouter un bail ou un avenant"
                    style={{
                      position: 'fixed', bottom: '28px', right: '32px', zIndex: 50,
                      display: 'flex', alignItems: 'center', gap: '8px',
                      padding: '13px 20px', borderRadius: '999px', border: 'none',
                      background: 'var(--accent)', color: '#fff', fontSize: '14px', fontWeight: 600,
                      cursor: 'pointer', boxShadow: '0 6px 20px rgba(26,95,168,.35)',
                    }}>
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
                    </svg>
                    Ajouter un bail / avenant
                  </button>
                )}

                {/* Modale d'ajout — regroupe dépôt, détection et extraction */}
                {showAddModal && (
                  <div className="modal-overlay" onClick={() => {
                    if (loading || detecting || compressing) return // fermeture bloquée pendant traitement
                    setShowAddModal(false)
                  }}>
                    <div className="modal" style={{ width: '95vw', height: '95vh', maxWidth: 'none', maxHeight: '95vh' }} onClick={e => e.stopPropagation()}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <div className="modal-title">Ajouter un bail ou un avenant</div>
                        <button
                          onClick={() => { if (!(loading || detecting || compressing)) setShowAddModal(false) }}
                          disabled={loading || detecting || !!compressing}
                          title={loading || detecting || compressing ? 'Traitement en cours…' : 'Fermer'}
                          style={{
                            background: 'none', border: 'none', fontSize: '20px', lineHeight: 1,
                            cursor: (loading || detecting || compressing) ? 'not-allowed' : 'pointer',
                            color: (loading || detecting || compressing) ? 'var(--text3)' : 'var(--text2)',
                            padding: '4px', opacity: (loading || detecting || compressing) ? 0.4 : 1,
                          }}>
                          ✕
                        </button>
                      </div>
                      <div style={{ overflowY: 'auto', flex: 1, paddingRight: '4px' }}>
                        <div className="extract-wrap">

                          {/* ── Queue principale ── */}
                          <>
                              <DropZone onFiles={handleFiles} disabled={loading || detecting || !!compressing} />
                    {compressing && (
                      <div className="warning-box" style={{ background: 'var(--accent-bg)', borderColor: 'rgba(26,95,168,.2)' }}>
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ flexShrink: 0, marginTop: '1px', color: 'var(--accent)' }}>
                          <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
                        </svg>
                        <span>
                          Compression de <strong>{compressing.name}</strong> (fichier volumineux)
                          {compressing.total > 0 ? ` — page ${compressing.current}/${compressing.total}` : '…'}
                        </span>
                      </div>
                    )}
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
                                    <span
                                      title={pertinent ? 'Cliquer pour marquer Non pertinent' : (raison ? `${raison} — Cliquer pour forcer Oui` : 'Cliquer pour forcer Oui')}
                                      onClick={() => {
                                        const newVal = !pertinent
                                        setPertinents(prev => { const n = [...prev]; n[fileIdx] = newVal; return n })
                                        // Si on passe à Oui et que le type n'est pas défini, mettre bail par défaut
                                        if (newVal && !docTypes[fileIdx]) {
                                          setDocTypes(prev => { const n = [...prev]; n[fileIdx] = 'bail'; return n })
                                        }
                                      }}
                                      style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '11px', fontWeight: 600,
                                        padding: '2px 8px', borderRadius: '999px', cursor: 'pointer',
                                        background: pertinent ? 'var(--success-bg)' : 'var(--danger-bg)',
                                        color: pertinent ? 'var(--success)' : 'var(--danger)',
                                        border: `1px solid ${pertinent ? 'var(--success)' : 'var(--danger)'}`,
                                        opacity: 0.9,
                                      }}>
                                      {pertinent ? 'Oui' : 'Non'}
                                      <span style={{ fontSize: '9px', opacity: 0.7 }}>⇄</span>
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
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '3px' }}>
                                      {avenantLinks[fileIdx]?.startsWith?.('dir-') && (
                                        <span style={{ fontSize: '10px', color: 'var(--success)', fontWeight: 600 }}>
                                          📁 Lié par répertoire
                                        </span>
                                      )}
                                      <select
                                        value={avenantLinks[fileIdx] || ''}
                                        onChange={e => setAvenantLinks(prev => ({ ...prev, [fileIdx]: e.target.value || null }))}
                                        style={{ fontSize: '11px', padding: '3px 6px', borderRadius: '6px', border: `1px solid ${avenantLinks[fileIdx]?.startsWith?.('dir-') ? 'var(--success)' : 'var(--border2)'}`, background: 'var(--surface)', color: avenantLinks[fileIdx] ? 'var(--text)' : 'var(--text3)', cursor: 'pointer', width: '100%' }}
                                      >
                                        <option value="">— Bail lié —</option>
                                        {/* Option virtuelle pour les liens dir- (avant extraction) */}
                                        {avenantLinks[fileIdx]?.startsWith?.('dir-') && (() => {
                                          const bailIdx = parseInt(avenantLinks[fileIdx].replace('dir-', ''))
                                          const bailFile = files[bailIdx]
                                          return bailFile ? (
                                            <option key={avenantLinks[fileIdx]} value={avenantLinks[fileIdx]}>
                                              {bailFile.name.replace(/\.[^.]+$/, '')}
                                            </option>
                                          ) : null
                                        })()}
                                        {allBails.map(b => (
                                          <option key={b.id} value={b.id}>
                                            {b.data?.immeuble || b.data?.adresse || b.file_name}
                                          </option>
                                        ))}
                                      </select>
                                    </div>
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
                      </div>
                    </div>
                  </div>
                )}
              </>
            )}
          </div>
        </main>
      </div>
    </>
  )
}
