const xlsx = require('xlsx')
const math = require('mathjs')

const workbook = xlsx.readFile('./Book2.xlsx')

const worksheet = workbook.Sheets['Sheet1']



const loads = xlsx.utils.sheet_to_json(worksheet)



let voltage_i = math.complex(11459.1664954711, 3.42112964173702);
for (let i = 0; i < loads.length; i++) {
    const z_i = math.complex(loads[i].R_OHM, loads[i].X_OHM);
    const current_i = math.complex(loads[i].Re_line_current, loads[i].Imz_line_current);
    const voltage_i_1 = math.subtract(voltage_i, math.multiply(z_i, current_i));

    voltage_i = voltage_i_1;
    loads[i].Re_Node_voltages_iteration_1 = math.re(voltage_i)
    loads[i].Imz_Node_voltages_iteration_1 = math.im(voltage_i)
    
  
}



const newWorkbook = xlsx.utils.book_new()
const newWorksheet = xlsx.utils.json_to_sheet(loads)
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1')

xlsx.writeFile(newWorkbook, 'Book3.xlsx')
