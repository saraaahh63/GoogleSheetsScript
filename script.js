function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Remove Rows', functionName: 'removeExcessRows'},
    {name: 'RemoveColumns', functionName: 'removeExcessColumns'},
    {name: 'Freeze Header', functionName: 'freezeHeader'},
    {name: 'Embed URLs', functionName: 'embedURL'}
  ];
  spreadsheet.addMenu('Conversion', menuItems);
}

//-----------------------------------------------------------------------------------------------------------------------------

/*Step 1: Run function removeExcessRows to remove unnecessary rows - you will be prompted with an input box, 
enter the rows you would like removed separated by commas
(NOTE: Make sure to remove any rows above the header row) */
  
function removeExcessRows() {
    var ui = SpreadsheetApp.getUi(); // Same variations.
    
    var result = ui.prompt(
      'Please enter the rows you would like deleted:',
      '(separate your entries with commas)',
      ui.ButtonSet.OK_CANCEL);
    
    var button = result.getSelectedButton();
    var text = result.getResponseText();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var array = text.split(",");
    
    if (button == ui.Button.OK) {
      // User clicked "OK".
      for (i = 0; i < array.length; i++) {
        item = parseInt(array[i]) - 1;
        if (item === 0) {
          item = 1
        }
        sheet.deleteRow(item);
      }
      ui.alert('They have been succesfully removed.');
    } else if (button == ui.Button.CANCEL) {
      // User clicked "Cancel".
      ui.alert('No rows were removed.');
    }
    
  }
  
/*Step 1.5: Run function removeExcessColumns to remove unnecessary columns, you must enter the Columns as numbers and not letters
Column A is 1, B is 2 and so on*/
  
function removeExcessColumns() {
    var ui = SpreadsheetApp.getUi(); // Same variations.
    
    var result = ui.prompt(
      'Please enter the columns you would like deleted:',
      '(separate your entries with commas and use CAPITAL letters)',
      ui.ButtonSet.OK_CANCEL);
    
    var button = result.getSelectedButton();
    var text = result.getResponseText();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var array = text.split(",");
    
    if (button == ui.Button.OK) {
      // User clicked "OK".
      var oldx = 0;
      for (num = 0; num < array.length; num += 1) {
        x = array[num];
        if (x == 'A') {
          x = 1;
          oldx++;
        } else if (x == 'B') {
          x = 2;
          oldx++;
        } else if (x == 'C') {
          x = 3;
          oldx++;
        } else if (x == "D") {
          x = 4;
          oldx++;
        } else if (x == "E") {
          x = 5;
          oldx++;
        } else if (x == "F") {
          x = 6
          oldx++;
        } else if (x == "G") {
          x = 7
          oldx++;
        }
        if (oldx > 1) {
          x = x - 1;
        }
        sheet.deleteColumn(x);
      }
      ui.alert('They have been succesfully removed.');
    } else if (button == ui.Button.CANCEL) {
      // User clicked "Cancel".
      ui.alert('No columns were removed.');
    }
    
  }
  
  
//Step 2: Run function freezeHeader to freeze the header row, unfreeze all other rows/columns manually
  
function freezeHeader() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    
    sheet.setFrozenRows(1)
  }
  
/*Step 3: Run function embedURL to combine name column and URL column, you will be prompted to enter the columns you want combined
Please use the column LETTER this time*/
  
function embedURL() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    
    var ui = SpreadsheetApp.getUi(); // Same variations.
    
    var result = ui.prompt(
      'Please enter the columns you would like combined:',
      '(type the name column followed by the URL column, separate your entries with commas and use CAPITAL letters)',
      ui.ButtonSet.OK_CANCEL);
    
    var button = result.getSelectedButton();
    var text = result.getResponseText();
    var array = text.split(",");
    
    if (button == ui.Button.OK) {
      // User clicked "OK".
      x = array[0];
      if (x == 'A') {
        x = 1;
        newNameLetter = 'B';
      } else if (x == 'B') {
        x = 2;
        newNameLetter = 'C';
      } else if (x == 'C') {
        x = 3;
        newNameLetter = 'D';
      } else if (x == "D") {
        x = 4;
        newNameLetter = 'E';
      } else if (x == "E") {
        x = 5;
        newNameLetter = 'F';
      }
      newUrlLetter = array[1];
      if (newUrlLetter == 'A') {
        newUrlLetter = 'B';
      } else if (newUrlLetter == 'B') {
        newUrlLetter = 'C';
      } else if (newUrlLetter == 'C') {
        newUrlLetter = 'D';
      } else if (newUrlLetter == "D") {
        newUrlLetter = 'E';
      } else if (newUrlLetter== "E") {
        newUrlLetter = 'F';
      }
      sheet.insertColumnBefore(x);
      sheet.getRange(array[0]+'2').setValue('=HYPERLINK('+newUrlLetter+'2,'+newNameLetter+'2)');
      var responsetitle = ui.prompt('What is the title of this combined column?', ui.ButtonSet.OK_CANCEL);
      var title = responsetitle.getResponseText();
      sheet.getRange(array[0]+'1').setValue(title);
      numofrows = sheet.getLastRow();
      for (num = 2; num <= numofrows; num += 1) {
        sheet.getRange(array[0]+num).setValue('=HYPERLINK('+newUrlLetter+num+','+newNameLetter+num+')');
      }
      ui.alert('They have been succesfully combined.');
      var range1 = sheet.getRange(newUrlLetter+'1');
      var range2 = sheet.getRange(newNameLetter+'1');
      sheet.hideColumn(range1);
      sheet.hideColumn(range2);
      ui.alert('The original columns have been hidden from view.');
    } else if (button == ui.Button.CANCEL) {
      // User clicked "Cancel".
      ui.alert('No columns were combined.');
    }
    
    sheet.clear({ formatOnly: true, contentsOnly: false });
  }
