const { merge, getInfo } = require('./index');
const { readJson } = require('./lib/file');
const path = require('path')

async function getCreds(creds) {       
    try {
        return await readJson(creds);        
    } catch (e) {
        throw new Error(`Unable to read credentials from ${creds}`);
    }
   
}

function stringifyInfo(result) {
    return [
        `Title:     ${result.doc.title}`,
        `Id:        ${result.doc.spreadsheetId}`,
        (result.sheet ?
        `Sheet:     ${result.sheet.title}` : null),
        //`Author:    ${result.doc.author.name} - ${result.doc.author.email}`,
        //`Updated:   ${result.doc.updated}`
    ].filter(s=>s !== null).join('\n');
    
}

function stringifyMergeDetails(result) {    
    const modified = result.modified.map(cell => `${cell.a1Address.padStart(6, ' ')}: ${cell.oldValue} -> ${cell.value}`);
    const added = result.added.map(row => JSON.stringify(row));
    const deleted = result.deleted.map(row => JSON.stringify(row));
    return [
        `-- Modified (${modified.length})--`,
        ...modified,
        `-- Added (${added.length}) --`,
        ...added,
        `-- Deleted (${deleted.length}) --`,
        ...deleted].join('\n');
}

function stringifyMergeSummary(result) {
    return [
        '-------- Merge summary --------',
        ['Modified', result.modified.length],
        ['Missing', result.missing.length],
        ['Added', result.added.length],
        ['Deleted', result.deleted.length]
    ].map(item => item.length === 2 ? `${item[0].padEnd(15, '.')}${item[1]}` : item)
}
require('yargs')
    .usage('Usage: $0 <cmd> [options]')
    .command('info <spreadsheetId>', 'get info about document', {}, async (argv) => {    
        const info = await getInfo(argv.spreadsheetId, await getCreds(argv.creds));
        console.log(stringifyInfo({ doc: info }));
    })
    .command('merge <file> <spreadsheetId>', 'merge csv file into spreadsheet', {
        's': {
            alias: 'silent',
            description: 'Produce no output'
        },
        'v': {
            type: 'boolean',
            alias: 'verbose',
            description: 'Be more talkative. ie: show change details'
        },
        'k': {            
            alias: 'keyField',
            description: 'Key field.  First column in sheet is used as default'
        },
        'sheet': {
            description: 'Index or name of sheet to work with. First sheet is used as default'
        },
        'removeRows': {
            type: 'boolean',
            description: 'Also remove rows from sheet not found in file'
        },
        'updateOnly': {
            type: "boolean",            
            description: 'Update only, do not add rows to sheet'
        }
    }, async (argv) => {
        const mergeConfig = {
            keyField: argv.keyField,
            csvFile: argv.file,
            spreadsheetId: argv.spreadsheetId,
            sheet: argv.sheet,
            removeRows: argv.removeRows,
            addMissing: argv.updateOnly !== true,
            creds: await getCreds(argv.creds)
        };
        
        const result = await merge(mergeConfig);
        if (argv.silent)
            return;
            
        const o = [
            '-------- Document info -------',
            stringifyInfo(result),
            '',
            (argv.verbose ?
                stringifyMergeDetails(result) :
                stringifyMergeSummary(result))
        ]
                    
        console.log(o
            .filter(item => !!item)
            .join('\n'));
    })
    .options({
        'creds': {
            default: 'service_account.json',
            description: 'Service Account credential file'
        }
    })
    .help('h')
    .alias('h','help')
    .argv;

