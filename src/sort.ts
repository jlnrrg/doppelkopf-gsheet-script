const dataLabel: string = 'data'
const templateLabel: string = 'template'

const sheetLabelList: string[] = [dataLabel, templateLabel]

function sortSheetsByName() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheets = spreadsheet.getSheets()

  const sheetsFiltered = sheets.filter((value) => {
    const sheetName = value.getName()

    return !sheetLabelList.some((e) => e == sheetName)
  })

  sheetsFiltered.sort((a, b) => a.getName().localeCompare(b.getName()))

  sheetsFiltered.map((value, index) => {
    spreadsheet.setActiveSheet(value)
    spreadsheet.moveActiveSheet(index + 3)
  })
}
