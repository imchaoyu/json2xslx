import * as XLSX from 'xlsx/xlsx.mjs';
import * as fs from 'fs';

XLSX.set_fs(fs);

const outputFile = './excel/单词.xlsx'
const inputFile = [
    {
        name: '雅思',
        file: './json/yasi.json',
    },
    {
        name: '专四八',
        file: './json/zhuan.json',
    }
]
let names = [];
let sheets={};

for (let i = 0; i < inputFile.length; i++) {
    const {name,file} = inputFile[i];
    // 读取文件
    const data = fs.readFileSync(file, 'utf8')

    // 转换JSON为对象
    const jsonData = JSON.parse(data)

    // 将JSON转换为Excel Sheet
    const jsonWorkSheet = XLSX.utils.json_to_sheet(jsonData)
    names.push(name)
    sheets[name]=jsonWorkSheet
    
}

// 将Sheet写入Excel文件
let workBook = {
    SheetNames: names,
    Sheets: sheets
};

// 将workBook写入文件
XLSX.writeFile(workBook, outputFile);
console.log('转换成功！')
// // 读取文件
// const data = fs.readFileSync(inputFile, 'utf8')

// // 转换JSON为对象
// const jsonData = JSON.parse(data)

// // 将JSON转换为Excel Sheet
// const jsonWorkSheet = XLSX.utils.json_to_sheet(jsonData)

// // const workbook = XLSX.utils.book_new();
// // XLSX.utils.book_append_sheet(workbook, worksheet, "Dates");

// // 将Sheet写入Excel文件
// let workBook = {
//     SheetNames: ['雅思'],
//     Sheets: {
//       '雅思': jsonWorkSheet,
//     }
//   };
// // XLSX.writeFile({
// //     SheetNames: ['jsonWorkSheet'],
// //     Sheets: {
// //         'jsonWorkSheet': jsonWorkSheet
// //     }
// // }, outputFile)
// // 将workBook写入文件
// XLSX.writeFile(workBook, outputFile);

