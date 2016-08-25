// Global variables

var scriptProperties = PropertiesService.getScriptProperties();

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('ESN')
      .addItem('Configuration', 'showPromptConfiguration')
      .addItem('Search ESNcard', 'showPromptSearch')
      .addToUi();
  
}

function showPromptConfiguration() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Setting things up',
      'Introduce the ID of the Spreadsheet you wish to fetch from',
      ui.ButtonSet.OK_CANCEL);
  
  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  
  if (button == ui.Button.OK) {
    scriptProperties.setProperty('DOCUMENT_ID', text);
    ui.alert('The ID has been save correctly');
  } 
}

function showPromptSearch() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Search the data related to an ESNcard and paste it in this document',
      'ESNcard number',
      ui.ButtonSet.OK_CANCEL);
  
  var activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
  
  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  
  if (button == ui.Button.OK) {
    search(text, activeCell);
  } 
}

function search(esnCardNumber, activeCell) {
  var documentID = scriptProperties.getProperty('DOCUMENT_ID');

  if(documentID != "" && documentID != null) {
    
    var sheetESNcard = SpreadsheetApp.openById(documentID).getSheetByName("Respuestas de formulario 1");
    var sheetSearch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hoja 1");
    
    var esnCards = sheetESNcard.getRange("A:Z").getValues();
    
    for (var i = 0; i< esnCards.length; i++) {
      
      if(esnCards[i][1] == esnCardNumber && esnCards[i][1] != undefined){
        
        // nº ESNcard
        sheetSearch.getRange(activeCell.getA1Notation()).setValue(esnCards[i][1]);
        // nombre
        sheetSearch.getRange(activeCell.getRow(),activeCell.getColumn() + 1).setValue(esnCards[i][2]);
        // apellido
        sheetSearch.getRange(activeCell.getRow(), activeCell.getColumn() + 2).setValue(esnCards[i][3]);
        // fecha de nacimiento
        sheetSearch.getRange(activeCell.getRow(), activeCell.getColumn() + 3).setValue(esnCards[i][4]);
        // telefono
        sheetSearch.getRange(activeCell.getRow(), activeCell.getColumn() + 4).setValue(esnCards[i][5]);
        // email
        sheetSearch.getRange(activeCell.getRow(), activeCell.getColumn() + 5).setValue(esnCards[i][6]);
        // nacionalidad
        sheetSearch.getRange(activeCell.getRow(), activeCell.getColumn() + 6).setValue(esnCards[i][7]);
        
        break;
      } 
    }
    if(activeCell.isBlank()){
      SpreadsheetApp.getUi().alert("There is no one registered with the given ESNcard number");
    }
  } else {
    SpreadsheetApp.getUi().alert("You must set up and ID first");
  }
}
