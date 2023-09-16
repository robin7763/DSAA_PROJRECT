const xlsx = require('xlsx')
const math = require('mathjs')

const workbook = xlsx.readFile('./Book1.xlsx')

const worksheet = workbook.Sheets['Sheet1']
// {
//     from_bus: 'Bhagirathi Bhawan',
//     to_bus: 'Residential Area-1',
//     R_OHM: 0.732,
//     X_OHM: 0.574,
//     realLoad: 60,
//     reactiveLoad: 20,
//     Re_Node_voltages: 11.5,
//     Imz_Node_voltages: 0
//   }

const loads = xlsx.utils.sheet_to_json(worksheet)



for (let i = 1; i < loads.length; i++) {

    loads[i].reactiveLoad_new = loads[i].reactiveLoad - (loads[i].reactiveLoad / 20.75)

}



const newWorkbook = xlsx.utils.book_new()
const newWorksheet = xlsx.utils.json_to_sheet(loads)
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1')

xlsx.writeFile(newWorkbook, 'Book2.xlsx')
