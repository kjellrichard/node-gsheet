const csv = require('./csv');
const { merge } = require('../lib/merge');

(async () => {
    var services = (await csv.readCsv('c:\\temp\\services2.csv', { separator: ',' })).filter(s=>s.production === 'true');            

    const merged = await merge({
        data: services,
        
        creds: require('../.cred/spreadsheet-test-4ce4a41fdfa9.json'),
        spreadsheetId: '12bM_EskLRrLP6EnlUdtBd_vFOSJNVe4i_agXf59eKBU',
        addMissing: true,
        keyField: 'name'
    })
})();