/* This script helps save time by inserting three rows into the Google Doc.
Point of Contact: Alvaro Gomez, Special Campuses Academic Technology Coach, 210-363-1577
*/


function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Update Table')
    .addItem('😀 Add rows to table', 'addRowToTable')
    .addToUi();
}

function addRowToTable() {

  var style1 = {};
  style1[DocumentApp.Attribute.BOLD] = true;

  var document = DocumentApp.getActiveDocument();
  var body = document.getBody();

  var tables = body.getTables();
  if (tables.length > 0) {
    var table = tables[0];
    
    // Inserts a blank row at the top of the table. This is the first agenda item.
    var newRow0 = table.insertTableRow(0);
    var cell1_0 = newRow0.insertTableCell(0);
    var cell2_0 = newRow0.insertTableCell(1);

    // Insert another new row at the top of the table. This is the weekly check-in.
    var newRow1 = table.insertTableRow(0);
    
    var cell1 = newRow1.insertTableCell(0);
    var cell2 = newRow1.insertTableCell(1);

    cell1.clear();

    // Add a bulleted list to cell1
    var bulletList = cell1.editAsText();
    // bulletList.appendText('Kim - \n').setAttributes(style1);
    // bulletList.appendText('Melanie - \n').setAttributes(style1);
    bulletList.appendText('Janet - \n').setAttributes(style1);
    bulletList.appendText('Terry - \n').setAttributes(style1);
    bulletList.appendText('Al - \n').setAttributes(style1);
    bulletList.appendText('Stacy - \n').setAttributes(style1);
    
    var newRow2 = table.insertTableRow(0).setAttributes(style1);
    
    // Add date to the cells
    var currentDate = new Date();
    
    var cell1_2 = newRow2.insertTableCell(0).setAttributes(style1);
    var cell2_2 = newRow2.insertTableCell(1);

    var formattedMonth = (currentDate.getMonth() + 1).toString();
    var formattedDay = currentDate.getDate().toString();

    cell1_2.setText(formattedMonth + '/' + formattedDay);

    // Set background color for the entire row
    setRowBackgroundColor(newRow2, '#FCE5CD');

  } else {
    // If no tables are found, notify the user
    DocumentApp.getUi().alert('No tables found in the document.');
  }
}

function setRowBackgroundColor(row, color) {
  for (var i = 0; i < row.getNumCells(); i++) {
    row.getCell(i).setBackgroundColor(color);
  }
}