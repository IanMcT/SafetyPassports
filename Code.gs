//Add safety passport sheets
//Ian McTavish
//Sept 13, 2019
//First sheet in spreadsheet called Class List - student names (lastname, first name) are in column A
//Second sheet - called Template - this is the sheet I want to duplicate.  The spot for the student name is in cell B6
function addPassports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var templatesheet = ss.getSheetByName("Template");
  
  Logger.log(ss.getName());
  Logger.log(sheet.getName());
  
  var counter = 1;
  var student = sheet.getRange(counter, 1);
  
  while(student.getValue() != ""){
    student = sheet.getRange(counter, 1);
    Logger.log(student.getValue());
    if(student.getValue() == "")
    {
      Logger.log("Empty string");
    }else{
      //duplicate sheet
      ss.setActiveSheet(templatesheet);
      ss.duplicateActiveSheet();
      ss.getActiveSheet().setName(student.getValue());
      ss.getActiveSheet().getRange(6,2).setValue(student.getValue());
      
    }
    counter++;
  }
}
