const tackenLabel: string = 'T'
const bockLabel: string = 'B'
const turnLabel: string = '#'
const soloLabel: string = 'S'

const labelList: string[] = [tackenLabel, bockLabel, turnLabel, soloLabel]

// start function on edit
function onEditTransformCell(e: GoogleAppsScript.Events.SheetsOnEdit) {
  console.log('Start of onEdit')
  const row = e.range.getRow()

  const bock = getBockRange(row)
  const tacken = getTackenRange(row)
  const previousCell = getPreviousRange(e.range)

  const bockA1 = bock.getA1Notation()
  const tackenA1 = tacken.getA1Notation()

  const previousCellA1 = previousCell?.getA1Notation() ?? null

  var formular: string = ''
  switch (e.value) {
    case 'w':
      formular = `=${
        previousCellA1 != null ? `${previousCellA1}+` : ''
      }(${tackenA1}*IF(LEN(${bockA1})>0;POW(2;(LEN(${bockA1})));1))`
      break
    case 'l':
      formular = `=${
        previousCellA1 != null ? `${previousCellA1}-` : ''
      }(${tackenA1}*IF(LEN(${bockA1})>0;POW(2;(LEN(${bockA1})));1))`
      break
    case 'ws':
      formular = `=${
        previousCellA1 != null ? `${previousCellA1}+` : ''
      }(${tackenA1}*IF(LEN(${bockA1})>0;POW(2;(LEN(${bockA1})));1)*3)`
      break
    case 'ls':
      formular = `=${
        previousCellA1 != null ? `${previousCellA1}-` : ''
      }(${tackenA1}*IF(LEN(${bockA1})>0;POW(2;(LEN(${bockA1})));1)*3)`
      break
    case 'n':
      formular = previousCellA1 != null ? `=${previousCellA1}` : ''
      break
  }
  console.log('Range: ' + JSON.stringify(e.range.getRow))
  console.log('Formular: ' + formular)
  if (formular.length > 0) {
    e.range.setFormula(formular)
    addTurnNumber(row)
  }
}

// adds the next turn number in the respective column if the cell is blank
function addTurnNumber(row: number) {
  const sheet = SpreadsheetApp.getActiveSheet()
  const column = getColumnByName(sheet, turnLabel)
  const cell = sheet.getRange(row, column)
  if (cell.isBlank()) {
    const previousCellA1 = getPreviousRange(cell)?.getA1Notation() ?? null
    const formular: string = `=${
      previousCellA1 != null ? `${previousCellA1}+` : ''
    }1`
    console.log(formular)
    cell.setFormula(formular)
  }
}

// gets the previous cell (by row) if it is a number or blank
function getPreviousRange(
  cell: GoogleAppsScript.Spreadsheet.Range,
): GoogleAppsScript.Spreadsheet.Range | null {
  const previousCell = cell
    .getSheet()
    .getRange(cell.getRow() - 1, cell.getColumn())

  // if the number is real
  if (Number.isFinite(previousCell.getValue()) || previousCell.isBlank()) {
    return previousCell
  } else {
    return null
  }
}

// gets the length of the bock field
function getBockRange(row: number): GoogleAppsScript.Spreadsheet.Range {
  const sheet = SpreadsheetApp.getActiveSheet()
  const column = getColumnByName(sheet, bockLabel)
  console.log('Bock Column: ' + String(column))
  return sheet.getRange(row, column)
}

// gets the value of the tacken field
function getTackenRange(row: number): GoogleAppsScript.Spreadsheet.Range {
  const sheet = SpreadsheetApp.getActiveSheet()
  const column = getColumnByName(sheet, tackenLabel)
  console.log('Tacken Column: ' + String(column))
  return sheet.getRange(row, column)
}

// gets the index of the column by name
function getColumnByName(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  name: string,
): number {
  const range = sheet.getRange(1, 1, 1, sheet.getMaxColumns())
  const values = range.getValues()

  for (const row in values) {
    for (const col in values[row]) {
      const lowerCaseValue = String(values[row][col]).toLowerCase()
      const lowerCaseName = String(name).toLowerCase()

      if (lowerCaseValue == lowerCaseName) {
        return parseInt(col) + 1
      }
    }
  }

  throw 'failed to get column by name'
}
