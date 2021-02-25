import * as Xls from 'excel4node'

const workbook = new Xls.Workbook()
// set the hidden property to true to hide the sheet
// LOOKUPS sheet contains the lookup values
const lookUpSheet = workbook.addWorksheet('LOOKUPS', { hidden: false })
lookUpSheet.cell(1, 1).string('Values')
lookUpSheet.cell(2, 1).string('A value')
lookUpSheet.cell(3, 1).string('B value')
lookUpSheet.cell(4, 1).string('C value')
lookUpSheet.cell(5, 1).string('D value')

const dataSheet = workbook.addWorksheet('Data')
dataSheet.cell(1, 1).string('Column A')
dataSheet.addDataValidation({
    type: 'list',
    allowBlank: false,
    showErrorMessage: true,
    errorStyle: 'stop',
    errorTitle: 'Invalid value',
    error: 'Please select a valid value from the list',
    // This is the range where the list validation will apply. You can set it to the whole column like 'A:A',
    // but in this case the first column also will contains this constraint.
    // As far as I know Excel does not support the Google Sheet like syntax 'A2:A'
    sqref: 'A2:A1000',
    formulas: ['=LOOKUPS!$A$2:$A$5'],
  });

  workbook.write('example.xlsx')

  console.log('Excel generated')