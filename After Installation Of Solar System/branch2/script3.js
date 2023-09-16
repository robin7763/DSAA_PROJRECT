const xlsx = require('xlsx')
const math = require('mathjs')

const workbook = xlsx.readFile('./Book4.xlsx')

const worksheet = workbook.Sheets['Sheet1']



const loads = xlsx.utils.sheet_to_json(worksheet)

// {

//     R_OHM: 0.896,
//     X_OHM: 0.7011,
//     Z_OHM: 1.1376982069072623,
//     realLoad: 420,
//     reactiveLoad: 200,
//     Re_Node_voltages: 11.55,
//     Imz_Node_voltages: 0,
//     Re_line_current: 36.36363636363636,
//     Imz_line_current: -17.316017316017316,

//     Re_Node_voltages_iteration_1: 5977.390476190477,
//     Imz_Node_voltages_iteration_1: -572.7055411255401
// }



const kva_i = math.complex(420, 200-(200/20.75));

const voltage_i = math.complex(loads[loads.length - 1].Re_Node_voltages_iteration_1 / 1000, loads[loads.length - 1].Imz_Node_voltages_iteration_1 / 1000);

let current_i = math.divide(kva_i, voltage_i);
current_i = current_i.conjugate()

loads[loads.length - 1].Re_line_current_iteration_1 = math.re(current_i)
loads[loads.length - 1].Imz_line_current_iteration_1 = math.im(current_i)



for (let i = loads.length - 1; i >= 2; i--) {

    const kva_i_1 = math.complex(loads[i - 1].realLoad, loads[i - 1].reactiveLoad_new);
    const voltage_i_1 = math.complex(loads[i - 1].Re_Node_voltages_iteration_1 / 1000, loads[i - 1].Imz_Node_voltages_iteration_1 / 1000);
    let current_i_1 = math.divide(kva_i_1, voltage_i_1);

    current_i_1 = current_i_1.conjugate()

    if (i == loads.length - 1) {

        loads[i - 1].Re_line_current_iteration_1 = math.re(current_i_1)
        loads[i - 1].Imz_line_current_iteration_1 = math.im(current_i_1)

    }

    let current_i_2 = math.add(current_i_1, current_i)


    loads[i - 2].Re_line_current_iteration_1 = math.re(current_i_2)
    loads[i - 2].Imz_line_current_iteration_1 = math.im(current_i_2)
    current_i = current_i_2


}



const newWorkbook = xlsx.utils.book_new()
const newWorksheet = xlsx.utils.json_to_sheet(loads)
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1')

xlsx.writeFile(newWorkbook, 'Book5.xlsx')
