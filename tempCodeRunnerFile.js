
const newWorkbook = xlsx.utils.book_new()
const newWorksheet = xlsx.utils.json_to_sheet(loads)
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'New Data')

xlsx.writeFile(newWorkbook, 'new Data File .xlsx')
