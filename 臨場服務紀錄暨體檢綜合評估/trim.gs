//去除程式擷取spreadsheet 資料後, cell 後面會有很多空白的問題

function trim() {
  var trimss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('database');
  var rangess = trimss.getDataRange();
  var valuesss = rangess.getValues();

  var trimedValues = valuesss.map(function(row) {
        return row.map(function(cell){
            return typeof cell === 'string' ? cell.trim() : cell;


        })



  })

  rangess.setValues(trimedValues);
}
