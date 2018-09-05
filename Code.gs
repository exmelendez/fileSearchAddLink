function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Test Links')
    .addItem('Create Links', 'showDialog')
    .addToUi();
}

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('index')
                        .setWidth(200)
                        .setHeight(100);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Insert State Test Links');
}

var gDriveEla = DriveApp.getFoldersByName("Coding")
                     .next()
                     .getFoldersByName("State Testing Links")
                     .next()
                     .getFoldersByName("ELA")
                     .next();

var gDriveMath = DriveApp.getFoldersByName("Coding")
                     .next()
                     .getFoldersByName("State Testing Links")
                     .next()
                     .getFoldersByName("MATH")
                     .next();

function createTestLink(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeSheet = ss.getSheetName();
  var currentRow = ss.getCurrentCell().getRow();
  var lastRow = ss.getLastRow();
  var currentCol = ss.getCurrentCell().getColumn();
  var previousCol = currentCol - 1;
  
  for(var i = currentRow; i < lastRow+1; i++){
    var cellContent = ss.getRange(i, previousCol).getValue();
    var cell = ss.getRange(i, currentCol);
    
    if (cellContent  === "" || cellContent === null) {
      cell.setValue("Cell Empty");
      cell.setFontColor("red");
      
    } else {
      
      var fileExists = filePresent(cellContent, activeSheet);
      
      if(fileExists) {
        var fileURL = getURL(cellContent, activeSheet);
        cell.setValue(fileURL);
    
      } else {
        cell.setValue("File Not Present");
        cell.setFontColor("red");
      }
    } 
  }
}

//Returns boolean if parameter passed combined with the .tif extension file type is found in the specified folder
function filePresent(cellContent, subject) {
  
  if(subject === "ELA"){
    return gDriveEla.getFilesByName(cellContent + ".tif")
                 .hasNext();
    
  } else if (subject === "MATH") {
        return gDriveMath.getFilesByName(cellContent + ".tif")
                        .hasNext();
  }
}

//Returns URL string of the parameter entered with the extension .tif
function getURL(cellContent, subject){
  
  if(subject === "ELA"){
    return gDriveEla.getFilesByName(cellContent + ".tif")
                 .next()
                 .getUrl();
    
  } else if (subject === "MATH") {
    return gDriveMath.getFilesByName(cellContent + ".tif")
                     .next()
                     .getUrl();
  }
}