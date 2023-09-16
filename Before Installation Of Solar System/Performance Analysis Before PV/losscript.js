const xlsx = require('xlsx')
const math = require('mathjs')

const workbook = xlsx.readFile('./Final Data.xlsx')

const worksheet = workbook.Sheets['Sheet1']

// {
//     from_bus: 'Bhagirathi Bhawan',
//     to_bus: 'Residential Area-1',
//     R_OHM: 0.732,
//     X_OHM: 0.574,
//     Re_Node_voltages: 11.5,
//     Imz_Node_voltages: 0,
//     realLoad: 60,
//     reactiveLoad: 20,
//     Re_line_current: 7.826086956521739,
//     Imz_line_current: -3.4782608695652173,
//     Re_Node_voltages_iteration_1: 10457.212000000001,
//     Imz_Node_voltages_iteration_1: -163.9200869565218,
//     Re_line_current_iteration_1: 8.544441146214396,
//     Imz_line_current_iteration_1: -3.95904812254761,
//     Re_Node_voltages_iteration_2: 10373.30773498858,
//     Imz_Node_voltages_iteration_2: -170.84354034513427,
//     Err_1_2: 84.18942863525679
//   }

const loads = xlsx.utils.sheet_to_json(worksheet)



for (let i = 0; i < loads.length; i++) {

    const current = math.complex(loads[i].Re_line_current_iteration_1, loads[i].Imz_line_current_iteration_1)
    const current_mag = current.toPolar().r

    loads[i].i2rLoss= current_mag * current_mag * loads[i].R_OHM
    loads[i].i2xLoss = current_mag * current_mag * loads[i].X_OHM
    
}



const newWorkbook = xlsx.utils.book_new()
const newWorksheet = xlsx.utils.json_to_sheet(loads)
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1')

xlsx.writeFile(newWorkbook, 'Book2.xlsx')
