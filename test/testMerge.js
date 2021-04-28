const { expect, should } = require('chai');
const { merge, getRows } = require('../lib/gsheet');
const config = {
    creds: require('../.cred/service_account.json'),
    spreadsheetId: '1dk8k4L30_2UJdvfkTw7bw2FD6RGaooVPLMTt7-CMdKw'
}
async function sleep(ms) {
    return new Promise(resolve => { 
        setTimeout(() => resolve(), ms);
    });
}


describe('merge', () => {
    it('it should remove all rows', async () => {   
        
        const res = await merge(Object.assign({}, config, { data: [], removeRows: true }));
        await sleep(5000);
        const rows = await getRows(config);        
        expect(rows).to.have.lengthOf(0);
    })

    it('it should add data', async () => {   
        const data = [
            { name: 'Peter', born: 1940, height: 180, gender: 'male' },
            { name: 'Jack', born: 1967, height: 176, gender: 'male' },
            { name: 'June', born: 1982, height: 160, gender: 'female' },
            { name: 'John', born: 1973, height: 190, gender: 'male' },
            { name: 'Mary', born: 1957, height: 168, gender: 'female' },
            { name: 'May', born: 1943, gender: 'female' }
        ];
        const res = await merge(Object.assign({}, config, { data: data }));
        await sleep(5000);
        const rows = await getRows(config); 
        const dataTotalBorn = data.reduce((acc, r) => acc += r.born, 0);
        const totalBorn = rows.reduce((acc, r) => acc += parseInt(r.born), 0);
        
        expect(rows).to.have.lengthOf(6);
        expect(dataTotalBorn).to.be.equal(totalBorn);        
    })
})

describe('file merge', () => {
    it('should read data from file', async () => {
        const r = await merge(Object.assign({}, config, { csvFile: __dirname + '/people.csv' }));
        expect(r).to.have.property('doc');
    })
})