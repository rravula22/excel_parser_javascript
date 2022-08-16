const axios = require('axios');
const Excel = require('exceljs');
const workbook = new Excel.Workbook();
const apb = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
let params = {
  access_key: 'e78a0264da0e76b7dbc3953b4a3d7cf2',
  query: ''
}
let filePath = process.argv.slice(2)[0]
workbook.xlsx.readFile(filePath)
.then(async function() {
    var worksheet = workbook.getWorksheet('Upated Jefferson_Shelby_YK AB');
    var masterWorksheet = workbook.getWorksheet('Instructions');
    var i=2;
    let keys = []
    headers = worksheet.getRow(1).values
    lastRec = headers.length
    addr = apb[headers.indexOf('Address') - 1];
    for( row in worksheet) {
      r=worksheet.getRow(i).values;
      if(r.indexOf('N/A') != -1) {
        keys = getAllOcc(r, 'N/A', i)
      }
      if(keys.length) {
        keys.forEach( (ele)=> {
          worksheet.getCell(ele).value = "";
        })
      }
      let corrdKey = addr + String(i)
      if(i == 20) {
        console.log("hi")
      }
      if(worksheet.getCell(corrdKey).value){
          params.query = worksheet.getCell(corrdKey).value
          await axios.get('http://api.positionstack.com/v1/forward', { params: params }).then((res)=>{
            if(res?.data?.data) {
              cdn = res.data.data[0];
              crd = apb[lastRec] + String(i)
              worksheet.getCell(crd).value = `${cdn.latitude}, ${cdn.longitude}`
            }
          })
      }
      i++;
    }

    return workbook.xlsx.writeFile('file.xlsx')
   });

function getAllOcc(row = [], val = 'N/A', index) {
    var keys = [], i = -1;
    while ((i = row.indexOf(val, i+1)) != -1){
        keys.push(apb[i -1] + String(index));
    }
    return keys;
}


// const reader = require('xlsx')
// file = reader.readFile("./jeff_data.xlsx")

callGeo = async (data) => {
    params.query = data;
    

}
