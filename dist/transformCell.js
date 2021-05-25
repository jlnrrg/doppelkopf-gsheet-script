"use strict";
const tackenLabel = 'T';
const bockLabel = 'B';
const turnLabel = '#';
const soloLabel = 'S';
const labelList = [tackenLabel, bockLabel, turnLabel, soloLabel];
// start function on edit
function onEditTransformCell(e) {
    var _a;
    console.log('Start of onEdit');
    const row = e.range.getRow();
    const bock = getBockRange(row);
    const tacken = getTackenRange(row);
    const previousCell = getPreviousRange(e.range);
    const bockA1 = bock.getA1Notation();
    const tackenA1 = tacken.getA1Notation();
    const previousCellA1 = (_a = previousCell === null || previousCell === void 0 ? void 0 : previousCell.getA1Notation()) !== null && _a !== void 0 ? _a : null;
    var formular = '';
    switch (e.value) {
        case 'w':
            formular = `=${previousCellA1 != null ? `${previousCellA1}+` : ''}(${tackenA1}*POW(2;${bockA1}))`;
            break;
        case 'l':
            formular = `=${previousCellA1 != null ? `${previousCellA1}-` : ''}(${tackenA1}*POW(2;${bockA1}))`;
            break;
        case 'ws':
            formular = `=${previousCellA1 != null ? `${previousCellA1}+` : ''}(${tackenA1}*POW(2;${bockA1})*3)`;
            break;
        case 'ls':
            formular = `=${previousCellA1 != null ? `${previousCellA1}-` : ''}(${tackenA1}*POW(2;${bockA1})*3)`;
            break;
    }
    console.log('Range: ' + JSON.stringify(e.range.getRow));
    console.log('Formular: ' + formular);
    if (formular.length > 0) {
        e.range.setFormula(formular);
        addTurnNumber(row);
    }
}
// adds the next turn number in the respective column if the cell is blank
function addTurnNumber(row) {
    var _a, _b;
    const sheet = SpreadsheetApp.getActiveSheet();
    const column = getColumnByName(sheet, turnLabel);
    const cell = sheet.getRange(row, column);
    if (cell.isBlank()) {
        const previousCellA1 = (_b = (_a = getPreviousRange(cell)) === null || _a === void 0 ? void 0 : _a.getA1Notation()) !== null && _b !== void 0 ? _b : null;
        const formular = `=${previousCellA1 != null ? `${previousCellA1}+` : ''}1`;
        console.log(formular);
        cell.setFormula(formular);
    }
}
// gets the previous cell (by row) if it is a number or blank
function getPreviousRange(cell) {
    const previousCell = cell
        .getSheet()
        .getRange(cell.getRow() - 1, cell.getColumn());
    // if the number is real
    if (Number.isFinite(previousCell.getValue()) || previousCell.isBlank()) {
        return previousCell;
    }
    else {
        return null;
    }
}
// gets the length of the bock field
function getBockRange(row) {
    const sheet = SpreadsheetApp.getActiveSheet();
    const column = getColumnByName(sheet, bockLabel);
    console.log('Bock Column: ' + String(column));
    return sheet.getRange(row, column);
}
// gets the value of the tacken field
function getTackenRange(row) {
    const sheet = SpreadsheetApp.getActiveSheet();
    const column = getColumnByName(sheet, tackenLabel);
    console.log('Tacken Column: ' + String(column));
    return sheet.getRange(row, column);
}
// gets the index of the column by name
function getColumnByName(sheet, name) {
    const range = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
    const values = range.getValues();
    for (const row in values) {
        for (const col in values[row]) {
            const lowerCaseValue = String(values[row][col]).toLowerCase();
            const lowerCaseName = String(name).toLowerCase();
            if (lowerCaseValue == lowerCaseName) {
                return parseInt(col) + 1;
            }
        }
    }
    throw 'failed to get column by name';
}
