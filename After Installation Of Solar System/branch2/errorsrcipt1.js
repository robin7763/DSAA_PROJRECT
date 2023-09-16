const xlsx = require('xlsx')
const math = require('mathjs')

const workbook = xlsx.readFile('./Book6.xlsx')

const worksheet = workbook.Sheets['Sheet1']




const loads = xlsx.utils.sheet_to_json(worksheet)
// {
//     from_bus: 'Bhagirathi Bhawan',
//     to_bus: 'Residential Area-1',
//     R_OHM: 0.732,
//     X_OHM: 0.574,
//     Re_Node_voltages: 11.5,
//     Imz_Node_voltages: 0,
//     Re_line_current: 7.826086956521739,
//     Imz_line_current: -3.4782608695652173,
//     Re_Node_voltages_iteration_1: 10457.212000000001,
//     Imz_Node_voltages_iteration_1: -163.9200869565218,
//     Re_line_current_iteration_1: 8.544441146214396,
//     Imz_line_current_iteration_1: -3.95904812254761,
//     Re_Node_voltages_iteration_2: 10373.30773498858,
//     Imz_Node_voltages_iteration_2: -170.84354034513427,
//     realLoad: 60,
//     reactiveLoad: 20
//   }

let max =0

for (let i = 0; i < loads.length; i++) {
   
    const v_1 = math.complex(loads[i].Re_Node_voltages_iteration_1, loads[i].Imz_Node_voltages_iteration_1)
    const v_2 = math.complex(loads[i].Re_Node_voltages_iteration_2, loads[i].Imz_Node_voltages_iteration_2)
    loads[i].Err_1_2 = Math.abs(math.subtract(v_1,v_2).toPolar().r)
    max = Math.max(max, math.subtract(v_1, v_2).toPolar().r)
  
}




const newWorkbook = xlsx.utils.book_new()
const newWorksheet = xlsx.utils.json_to_sheet(loads)
xlsx.utils.book_append_sheet(newWorkbook,newWorksheet,'Sheet1')

xlsx.writeFile(newWorkbook, 'Book7.xlsx')
