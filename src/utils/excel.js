import * as XLSX from "xlsx"

export function parseExcel(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" })

  return workbook.SheetNames.map(name => {
    const sheet = workbook.Sheets[name]

    const data = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: ""
    })

    return {
      name,
      data
    }
  })
}
