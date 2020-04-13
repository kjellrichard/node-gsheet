const { GoogleSpreadsheet } = require('google-spreadsheet');
const { readCsv } = require('./file');

function getSheetFromInfo(info, titleOrIndex=0) {
    if (typeof titleOrIndex === 'number')
        return info.sheetsByIndex[titleOrIndex];
	return info.sheetsByIndex.find(sheet => sheet.title.toLowerCase() === titleOrIndex.toLowerCase() )
} 

function keyify(value) {
    return value.toLowerCase().replace(/[\s\\\/_]/ig, '');
}

async function getHeaderColumns(sheet, headerRow) {    
    await sheet.loadCells({ 'startRowIndex': headerRow, 'endRowIndex': headerRow + 1})    
    const cells = [];
    for (let i = 0; i < sheet.columnCount; i++) {
        const cell = sheet.getCell(headerRow, i);
        cells.push({
            index: cell.columnIndex,
            col: cell.columnIndex,
            name: cell.value || '',
            key: keyify(cell.value || '')
        })
    }
    return cells;    
}

async function getCells(worksheet, startRow) {
    await worksheet.loadCells({ 'startRowIndex': startRow });            
    const cells = [];
    for (let r = startRow; r < worksheet.rowCount;r ++)  {
        for (let c = 0; c < worksheet.columnCount; c++) {
            cells.push(worksheet.getCell(r, c));
    }}
    return cells;
}

function findColumn(columns, value) {
    const toFind = keyify(value);
    return columns.filter(c => c.key === toFind)[0];
}

function createKeyMap(keyCells, columns, keyField) {
    const col = findColumn(columns, keyField);
    
    return keyCells
        .filter(c => c.columnIndex === col.col && !isEmpty(c.value) )
        .map(c=>({
            value: c.value,
            formattedValue: c.formattedValue,
            row: c.a1Row,
            rowIndex: c.rowIndex,
            keyField: keyify(keyField)
        }))
}

async function addRows(sheet, items) {
    await sheet.addRows(items);
}

function isEmpty(value) {
    return value === null || value === undefined || value === '';
}

async function createDoc(spreadsheetId, creds) {
    const doc = new GoogleSpreadsheet(spreadsheetId);
    if ( creds )
        await doc.useServiceAccountAuth(creds);    
    return doc;
}

async function getInfo(spreadsheetId, creds, sheet) {
    const doc = await createDoc(spreadsheetId, creds);    
    await doc.loadInfo();
    return doc;
}

async function getSheet(spreadsheetId, creds, sheet) {
    return getSheetFromInfo(await getInfo(spreadsheetId,creds), sheet); 
}

async function removeSheetRows(sheet, keyMap, allKeys) {
    const keysToRemove = keyMap.filter(key => {
        return allKeys.filter(k => k === key.value).length === 0;
    })
    if (!keysToRemove.length)
        return;
    const deletedRows = [];
    const rows = await sheet.getRows();
    for (let keyToRemove of keysToRemove) {
        const foundRow = rows.filter(row => row.rowIndex === keyToRemove.row)[0];        
        if (!foundRow)
            continue;
        deletedRows.push(foundRow);
        await foundRow.delete();
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
    return currentSheet.getRows(...arguments);
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
 * @param {Number} config.headerRow Index of where to find header row. Default: 0
 * @param {String} config.keyField Field to use as identifier. 
 * @param {Boolean} config.addMissing Add to spreadsheet if not already exists. Default: true. 
 * @param {Boolean} config.removeRows Remove rows in sheet not found in data. Default: false. 
 */
async function merge({ creds, spreadsheetId, sheet = 0, data, csvFile, headerRow=0, keyField, addMissing=true, removeRows=false }) {    
    if (csvFile)
        data = await readCsv(csvFile, { separator: undefined });   
    const doc = await getInfo(spreadsheetId, creds);
    const currentSheet = await getSheetFromInfo(doc, sheet);
    const columns = await getHeaderColumns(currentSheet, headerRow); 
    
    const cells = await getCells(currentSheet, headerRow + 1);
    if (!keyField || typeof keyField === 'number' || keyField == '0')
        keyField = columns[parseInt(keyField || 0)].name;
    const keyMap = createKeyMap(cells, columns, keyField);
    const missing = [];
    const modified = [];  
    let deleted = [];
    data.forEach(item => {
        const matchedKey = keyMap.filter(k => k.value == item[keyField] || k.formattedValue == item[keyField])[0];
        if (!matchedKey) {
            missing.push(item)
            return;
        }
        const matchedCells = cells.filter(cell => cell.rowIndex == matchedKey.rowIndex); 
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
            cell.oldValue = cell.value;
            cell.value = value;
            
            modified.push(cell);            
        });
    });

    if (missing.length > 0 && addMissing)
        await addRows(currentSheet, missing) 
    if (removeRows) {
        deleted = await removeSheetRows(currentSheet, keyMap, data.map(item => item[keyField]));
            
    }
    await currentSheet.saveUpdatedCells();
    //await currentSheet.bulkUpdateCells(modified);    
    return {
        doc: doc,
        sheet: currentSheet,
        modified,
        missing,
        added: addMissing ? missing : [],
        deleted
    }
}

exports.merge = merge;
exports.getSheet = getSheet;
exports.getRows = getRows;
exports.getInfo = getInfo;