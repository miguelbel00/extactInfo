var XLSX = require("xlsx");

const fs = require("fs")
let files = fs.readdirSync('./')
let filesNames = files.filter((f)=> f.includes('.xlsx') && !f.includes('new'))

try {
  var filenew = XLSX.readFile("./new.xlsx");
} catch (error) {
  console.log("Not Found File 'new.xlsx'")
  console.log("Creating File....")
  var filenew = XLSX.utils.book_new()
  
}

let sheets = [];

for (let i = 0; i < filesNames.length; i++) {
  let fileReaded = XLSX.readFile(filesNames[i]);
  sheets.push(
    XLSX.utils.sheet_to_json(fileReaded.Sheets[fileReaded.SheetNames[1]])
  );
}
let info = []

sheets.forEach((sheet, i) => {

  if (i == 0) {
    sheet.forEach((res) => {
        info.push({
          cuenta: res.Cuenta,
          nombre: res.Nombre,
          movimiento: res.Movimiento,
          saldoCierre: Object.values(res)[Object.values(res).length - 1],
        });
      
    });
  }else {
    sheet.forEach((res) => {
     if (i==2) {
     }
      let temp = {};
      temp[`movimiento${i}`]=res.Movimiento
      temp[`saldoCierre${i}`]=Object.values(res)[Object.values(res).length - 1]

      let found = info.find((o) => o.cuenta == res.Cuenta);
      if (found) {
        for (const key in temp) {
          found[key] = temp[key];
        }
      } else {
        
        temp.cuenta=res.Cuenta
        temp.nombre=res.Nombre
        info.push(temp);
      }
      
    });
  }
});

let fulldata = [...info];

fulldata.sort((a, b) => {
  return a.cuenta.toString() < b.cuenta.toString()
    ? -1
    : a.cuenta.toString() > b.cuenta.toString()
    ? 1
    : 0;
});

const workSheets = XLSX.utils.json_to_sheet(fulldata);
XLSX.utils.book_append_sheet(filenew, workSheets);
XLSX.writeFile(filenew, "./new.xlsx"); 
console.log("Completed")


