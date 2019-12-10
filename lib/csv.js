const {promisify} = require('util');
const fs = require('fs');
const readFile = promisify(fs.readFile); 
const writeFile = promisify(fs.writeFile);

function detectSeparator(line) {
    for (let s of [',', ';', '\t']) {
        if (line.indexOf(s) > -1)
            return s;
    }
    throw new Error('Separator could not be detected');
}
async function readCsv(fileName, { separator = undefined, iterator = null, skipFields = [] }) {
    let r = await readFile(fileName, 'utf8');
    if (r.charCodeAt(0) === 0xFEFF) {
        r = r.substr(1);
    }

    let headers;
    const content = r.split('\n').map((l, i) => {   
        if (i === 0 && separator === undefined) {
            separator = detectSeparator(l);
        }
        const parts = l.replace(/\r|\n|/ig, '').split(separator);   
        if (i === 0) {            
            headers = parts;                                
        }
        const info = {};
        headers.forEach((v, i) => {
            if (skipFields.indexOf(v) !== -1)
                return;            
            info[v] = parts[i];

        })
        if (iterator)
            iterator(info);
        return info;
    })
    return content.slice(1);
}

async function writeCsv(collection, filename, separator = ';', skipFields = []) {
    const fields = Object.keys(collection[0]).filter(v=>skipFields.indexOf(v)===-1);
    const lines = collection.map(member => {
        return fields.reduce((acc, value) => {
            acc.push(member[value]);
            return acc;
        }, []).join(separator)
    });
    await writeFile(filename, [fields.join(separator),  ...lines].join('\n'));
}

exports.readCsv = readCsv;
exports.writeCsv = writeCsv;