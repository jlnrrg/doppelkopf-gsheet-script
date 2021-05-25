function getUserNames(sheet: GoogleAppsScript.Spreadsheet.Sheet): string[] {
  var range = sheet.getRange(1, 1, 1, sheet.getMaxColumns())
  var values = range.getValues()

  var list: string[] = []

  for (var row in values) {
    for (var col in values[row]) {
      let value = String(values[row][col])

      // value is not part of labelList and not empty
      if (!labelList.some((e) => e == value) && value.length > 0) {
        let firstName = value.split(' ')[0]

        list.push(firstName)
      }
    }
  }
  return list.sort()
}
