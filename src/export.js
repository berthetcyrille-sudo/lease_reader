import * as XLSX from 'xlsx'
import { ALL_FIELDS, SECTIONS } from './fields.js'

export function exportToExcel(data, fileName) {
  const wb = XLSX.utils.book_new()

  // --- Onglet 1 : Tableau (1 ligne par bail) ---
  const headers = ALL_FIELDS.map(f => f.label)
  const values  = ALL_FIELDS.map(f => data[f.key] ?? '')
  const wsTable = XLSX.utils.aoa_to_sheet([headers, values])
  wsTable['!cols'] = headers.map(() => ({ wch: 24 }))
  // Style entête (fond bleu, texte blanc) via commentaire — XLSX ne supporte pas
  // les styles complets sans xlsx-style ; on garde simple
  XLSX.utils.book_append_sheet(wb, wsTable, 'Tableau')

  // --- Onglet 2 : Fiche (clé / valeur) avec sections ---
  const ficheRows = [['Section', 'Champ', 'Valeur']]
  SECTIONS.forEach(sec => {
    sec.fields.forEach(f => {
      ficheRows.push([sec.label, f.label, data[f.key] ?? ''])
    })
  })
  const wsFiche = XLSX.utils.aoa_to_sheet(ficheRows)
  wsFiche['!cols'] = [{ wch: 26 }, { wch: 32 }, { wch: 64 }]
  XLSX.utils.book_append_sheet(wb, wsFiche, 'Fiche')

  const baseName = fileName.replace(/\.[^.]+$/, '') || 'bail'
  XLSX.writeFile(wb, `lease_abstract_${baseName}.xlsx`)
}
