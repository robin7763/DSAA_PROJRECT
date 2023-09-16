const xlsx = require('xlsx')
const math = require('mathjs')

const workbook = xlsx.readFile('./Book2.xlsx')

const worksheet = workbook.Sheets['Sheet1']


// R_OHM: 0.896,
//     X_OHM: 0.7011,
//         Z_OHM: 1.1376982069072623,
//             realLoad: 420,
//                 reactiveLoad: 200,
//                     Re_Node_voltages: 11.55,
//                         Imz_Node_voltages: 0,
//                             Re_line_current: 36.36363636363636,
//                                 Imz_line_current: -17.316017316017316

const loads = xlsx.utils.sheet_to_json(worksheet)


let max = 0
let voltage_i = math.complex(11500, 0);
for (let i = 0; i < loads.length; i++) {
    const z_i = math.complex(loads[i].R_OHM, loads[i].X_OHM);
    const current_i = math.complex(loads[i].Re_line_current, loads[i].Imz_line_current);
    const voltage_i_1 = math.subtract(voltage_i, math.multiply(z_i, current_i));

    voltage_i = voltage_i_1;
    loads[i].Re_Node_voltages_iteration_1 = math.re(voltage_i)
    loads[i].Imz_Node_voltages_iteration_1 = math.im(voltage_i)
    max = Math.max(max, voltage_i.toPolar().r)
  
}
console.log(max)


// const newWorkbook = xlsx.utils.book_new()
// const newWorksheet = xlsx.utils.json_to_sheet(loads)
// xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1')

// xlsx.writeFile(newWorkbook, 'Book3.xlsx')
