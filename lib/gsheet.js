const GoogleSpreadsheet = require('google-spreadsheet');
const { promisify } = require('util');
const { readCsv } = require('./csv');
function promisifyAll(obj) {
    Object.keys(obj).forEach(key => {
        if (typeof obj[key] !== 'function' )
            return;
        const f = promisify(obj[key]).bind(obj);
        obj[key + 'Async'] = f;
    })
    return obj;
}

function getSheetFromInfo(info, titleOrIndex=0) {
        
    if (typeof titleOrIndex === 'number')
        return promisifyAll(info.worksheets[titleOrIndex]);
    var s = info.worksheets.filter(sheet => sheet.title.toLowerCase() === titleOrIndex)[0];
    return s ? promisifyAll(s) : null;    
} 

function keyify(value) {
    return value.toLowerCase().replace(/[\s\\\/_]/ig, '');
}

async function getHeaderColumns(sheet, headerRow) {
    const cells = await sheet.getCellsAsync({ 'min-row': headerRow, 'max-row': headerRow, 'return-empty': true })
    return cells.map((cell, index) => ({
        index: index,
        col: cell.col,
        name: cell.value || '',
        key: keyify(cell.value || '')
    }));    
}

function findColumn(columns, value) {
    const toFind = keyify(value);
    return columns.filter(c => c.key === toFind)[0];
}

async function collectKeys(currentSheet, columns, keyField) {
    const col = findColumn(columns, keyField);
    const cells = await currentSheet.getCellsAsync({ 'min-row': 1, 'max-row': currentSheet.rowCount, 'return-empty': true, 'min-col': col.col, 'max-col': col.col });
    return cells;
}

function createKeyMap(keyCells,columns, keyField) {
    const col = findColumn(columns, keyField);
    return keyCells
        .filter(c => c.col === col.col && !isEmpty(c.value) )
        .map(c=>({
            value: c.value,
            row: c.row,
            keyField: keyify(keyField)
        }))
}

async function addRows(sheet, items) {
    for (let item of items) {
        await sheet.addRowAsync(item);
    }
}

function isEmpty(value) {
    return value === null || value === undefined || value === '';
}

async function createDoc(spreadsheetId, creds) {
    const doc = promisifyAll(new GoogleSpreadsheet(spreadsheetId));
    if ( creds )
        await doc.useServiceAccountAuthAsync(creds);    
    return doc;
}

async function getSheet(spreadsheetId, creds, sheet) {
    const doc = await createDoc(spreadsheetId, creds);    
    const info = await doc.getInfoAsync();
    return getSheetFromInfo(info, sheet); 
}

async function removeSheetRows(sheet, keyMap, allKeys) {
    const keysToRemove = keyMap.filter(key => {
        return allKeys.filter(k => k === key.value).length === 0;
    })
    if (!keysToRemove.length)
        return;
    const deletedRows = [];
    for (let keyToRemove of keysToRemove) {
        const rows = await sheet.getRowsAsync({ query: `${keyToRemove.keyField} = ${keyToRemove.value}` });
        if (rows[0]) {
            deletedRows.push(rows[0]);
            await promisify(rows[0].del)();
        }
    }
    return deletedRows;
}

/**
 * 
 * @param {Object} config getRows config
 * @param {Object} config.creds Google Service Account credentials
 * @param {String} config.spreadsheetId Google Spreadsheet Id / key. The long id in the url
 * @param {String|Number} config.sheet Title or index of spreadsheet. Default: 0
 */
async function getRows({spreadsheetId,creds,sheet}) {
    const currentSheet = await getSheet(spreadsheetId,creds,sheet); 
    return currentSheet.getRowsAsync(...arguments);
}
/**
 * Merge data into an existing spreadsheet. Overwrites spreadsheet value with values from data. 
 * NOTE: Uses cell-based approach for matching. Row-based is fine, but has no bulk update. 
 * @param {Object} config merge config
 * @param {Object} config.creds Google Service Account credentials
 * @param {String} config.spreadsheetId Google Spreadsheet Id / key. The long id in the url
 * @param {String|Number} config.sheet Title or index of spreadsheet. Default: 0
 * @param {Array} config.data Items to merge into spreadsheet
 * @param {String} config.csvFile Filepath to read data from (overrides `config.data`)
 * @param {Number} config.headerRow Index of where to find header row. Default: 1
 * @param {String} config.keyField Field to use as identifier. 
 * @param {Boolean} config.addMissing Add to spreadsheet if not already exists. Default: true. 
 * @param {Boolean} config.removeRows Remove rows in sheet not found in data. Default: false. 
 */
async function merge({ creds, spreadsheetId, sheet = 0, data, csvFile, headerRow=1, keyField, addMissing=true, removeRows=false }) {    
    if (csvFile)
        data = await readCsv(csvFile, {separator: undefined});
    const currentSheet = await getSheet(spreadsheetId, creds, sheet);
    const columns = await getHeaderColumns(currentSheet, headerRow);    
    const cells = await currentSheet.getCellsAsync({ 'min-row': headerRow + 1, 'max-row': currentSheet.rowCount, 'return-empty': true });            
    const keyMap = createKeyMap(cells, columns, keyField);
    const missing = [];
    const modified = [];  
    let deleted = [];
    data.forEach(item => {
        const matchedKey = keyMap.filter(k => k.value == item[keyField])[0];
        if (!matchedKey) {
            missing.push(item)
            return;
        }
        const matchedCells = cells.filter(cell => cell.row == matchedKey.row); 
        
        Object.keys(item).forEach(key => {            
            const value = item[key];
            if (isEmpty(value))
                return;
            const col = findColumn(columns, keyify(key));
            if (!col)
                return;
            const cell = matchedCells[col.index];
            if (!cell || cell.value == value)
                return;           
            cell.value = value;           
            modified.push(cell);            
        });
    });

    if (missing.length > 0 && addMissing)
        addRows(currentSheet, missing) 
    if (removeRows) {
        deleted = await removeSheetRows(currentSheet, keyMap, data.map(item => item[keyField]));
            
    }
    await currentSheet.bulkUpdateCellsAsync(modified);    
    return {
        modified,
        missing,
        added: addMissing ? missing : [],
        deleted
    }
}

exports.merge = merge;
exports.getSheet = getSheet;
exports.getRows = getRows;