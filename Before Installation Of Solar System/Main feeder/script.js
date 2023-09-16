const xlsx = require('xlsx')
const math = require('mathjs')

const workbook = xlsx.readFile('./Book1.xlsx')

const worksheet = workbook.Sheets['Sheet1']



const loads = xlsx.utils.sheet_to_json(worksheet)

//residential load 
const kva_i = math.complex(90,40);

const voltage_i = math.complex(11.5,0);

let current_i = math.divide(kva_i, voltage_i);

current_i=current_i.conjugate()


loads[loads.length - 1].Re_line_current = math.re(current_i)
loads[loads.length - 1].Imz_line_current = math.im(current_i) 



for (let i = loads.length - 1; i >= 2; i--) {

  let kva_i_1 = math.complex(loads[i - 1].realLoad, loads[i - 1].reactiveLoad);
  let voltage_i_1 = math.complex(loads[i - 1].Re_Node_voltages, loads[i - 1].Imz_Node_voltages);
  let current_i_1 = math.divide(kva_i_1, voltage_i_1);

  current_i_1=current_i_1.conjugate()

  if (i == loads.length - 1) {
   
    loads[i-1].Re_line_current = math.re(current_i_1)
    loads[i - 1].Imz_line_current = math.im(current_i_1) 

  }
 
  let current_i_2 = math.add(current_i_1, current_i)
  
  
  loads[i-2].Re_line_current = math.re(current_i_2)
  loads[i-2].Imz_line_current = math.im(current_i_2)
  current_i = current_i_2

 
}



const newWorkbook = xlsx.utils.book_new()
const newWorksheet = xlsx.utils.json_to_sheet(loads)
xlsx.utils.book_append_sheet(newWorkbook,newWorksheet,'Sheet1')

xlsx.writeFile(newWorkbook, 'Book2.xlsx')
