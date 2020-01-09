("use strict");
// import XLSX from "xlsx";
const XLSX = require("js-xlsx");
const fs = require("fs");
const file_name = "toubiao.xlsx";
// const workbook = XLSX.readFile(file_name, { encoding: "utf8" });
const file = fs.readFileSync(file_name, { encoding: "utf8" });
// console.log(file);
const workbook = XLSX.readFile(file, { type: "array" });

// const sheet_name_list = workbook.SheetNames;
// var first_sheet_name = workbook.SheetNames[0];
// var worksheet = workbook.Sheets[first_sheet_name];
// var address_of_cell = "C3";
// var desired_cell = worksheet[address_of_cell];
// var desired_value = desired_cell.v;
// console.log(desired_cell);
// const sheetList = sheet_name_list.map((sheetName) => {
//     let workSheet = workbook.Sheets[sheetName];
//     console.log(workSheet);
//     fs.writeFile("./aaa.txt", JSON.stringify(workSheet, null, 4), "utf8", function(err) {
//         if (err) return console.log(err);
//     });
//     return XLSX.utils.sheet_to_json(workSheet);
// });
// console.log(sheetList);
