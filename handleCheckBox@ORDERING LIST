function handleCheckboxEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  //CONFIGURATION
  var sheetName = "ORDERING LIST";
  var checkboxColumn = 7; // Column G
  var dateColumn = 8;     // Column H
  var releaseTypeCol = 9; // Column I
  var secretPassword = "123";
  // ---------------------

  //1. VALIDATION: Check if we are on the right sheet and column
  if (sheet.getName() !== sheetName || range.getColumn() !== checkboxColumn || range.getRow() <= 1) {
    return;
  }

  //2. DETECT VALUE CHANGE
  //e.value is the NEW value (String). 
  //check range.getValue() to make sure is boolean status.
  var isChecked = range.getValue() === true;
  
  //USER CHECKED THE BOX (TRUE)
  //Action: Add Date Timestamp
  if (isChecked) {
    var timestamp = Utilities.formatDate(new Date(), e.source.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    sheet.getRange(range.getRow(), dateColumn).setValue(timestamp);
  }

  //USER UNCHECKED THE BOX (FALSE)
  //Action: Demand Password. If fail, revert to TRUE.
  else {
    //only ask for password if it was previously checked.
    //e.oldValue gives us the previous state
    
    //Prompt user
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Authentication Required', 'Please enter password to uncheck:', ui.ButtonSet.OK_CANCEL);
    
    //Process User Input
    if (response.getSelectedButton() == ui.Button.OK) {
      var inputPass = response.getResponseText();
      
      if (inputPass === secretPassword) {
        //PASSWORD CORRECT:
        //Allow the uncheck to happen.
        //Clear the Date (Col H) and Release Type (Col I) if desired
        sheet.getRange(range.getRow(), dateColumn).clearContent();
        
        //Reset Release Type to empty 
        sheet.getRange(range.getRow(), releaseTypeCol).clearContent(); 
        
        e.source.toast("Item Unchecked Successfully.", "Success");
      } else {
        //PASSWORD WRONG:
        e.source.toast("Incorrect Password. Action Reverted.", "Error");
        //Revert checkbox back to TRUE
        range.setValue(true);
      }
    } else {
      //USER CLICKED CANCEL:
      e.source.toast("Action Cancelled.", "Cancelled");
      //Revert checkbox back to TRUE
      range.setValue(true);
    }
  }
}
