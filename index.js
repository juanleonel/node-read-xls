const reader = require('xlsx')
const fs = require('fs');

const PATH = 'C:\\CPdescarga.xlsx'

function buildObject(row, sheet_name) {
  return {
    _sheet_name: sheet_name,
    postalCode: row.d_codigo,
    state: row.d_estado,
    municipality: row.D_mnpio,
    settlement: row.d_asenta
  }
}

function writeFiles(fileName, dataContent) {
  return new Promise((resolve, reject) =>{
    fs.writeFile("data/" + fileName + ".json", dataContent, 'utf8', function (err) {
      if (err) {
        reject(false)
      }
      resolve(true);
    });
  })
}

function main() {
  const file = reader.readFile(PATH)
  let data = []
  const sheets = file.SheetNames.filter(sheet => sheet !== 'Nota')
  for(let i = 0; i < sheets.length; i++) {
    const _sheet = reader.utils.sheet_to_json(file.Sheets[sheets[i]])
    if (_sheet.length) {
      for (const row of _sheet) {
        data.push(buildObject(row, sheets[i]))
      }
    }
  }

  for (const _sheet of sheets) {
    let objects =  data.filter(element => element._sheet_name === _sheet)
    objects = objects.map(element => {
      delete element._sheet_name
      return element
    })
    const jsonContent = JSON.stringify(objects);
    writeFiles(_sheet, jsonContent).then(result => {
      if (result) {
        console.log('json created');
      }
    }).catch(err =>{
      console.log(err);
    });
  }
}

main();