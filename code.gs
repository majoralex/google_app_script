function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Dialog Examples')
    .addItem('Show Alert', 'showAlert')
    .addItem('Show Input Box', 'showInputBox')
    .addToUi();
}

function showAlert() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Sample Alert',
    'This is a simple alert dialog box.',
    ui.ButtonSet.OK
  );
}

function showInputBox() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Input Example',
    'Please enter some text:',
    ui.ButtonSet.OK_CANCEL
  );

  // Get the button that was clicked
  const button = result.getSelectedButton();
  // Get the text field value
  const text = result.getResponseText();
  
  if (button === ui.Button.OK) {
    try {
      // Get the specific sheet
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('input_box_data');
      
      if (!sheet) {
        ui.alert('Error', 'Sheet "input_box_data" not found. Please create this sheet first.', ui.ButtonSet.OK);
        return;
      }
      
      // Get the next empty row in column A
      const lastRow = sheet.getLastRow();
      const nextRow = lastRow + 1;
      
      // Write the input data and timestamp
      sheet.getRange(nextRow, 1).setValue(text); // Column A
      sheet.getRange(nextRow, 2).setValue(new Date()); // Column B
      
      // Format timestamp column
      sheet.getRange(nextRow, 2).setNumberFormat('MM/dd/yyyy HH:mm:ss');
      
      ui.alert('Success', 'Data has been written to the input_box_data sheet!', ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('Error', 'Failed to write data: ' + error.toString(), ui.ButtonSet.OK);
    }
  }
}
