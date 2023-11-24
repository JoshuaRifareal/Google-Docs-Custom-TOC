function createCustomTOC() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  var tocTitle = "Custom TOC";
  var tableExists = false;
  var tables = body.getTables();

  var table;

  for (var i = 0; i < tables.length; i++) {
    var numRows = tables[i].getNumRows();
    var numCells = numRows > 0 ? tables[i].getRow(0).getNumCells() : 0;

    if (numRows > 0 && numCells > 0) {
      var titleCell = tables[i].getRow(0).getCell(0);
      if (titleCell.getText() === tocTitle) {
        tableExists = true;
        table = tables[i];
        // Clear existing rows in the table
        var rows = table.getNumRows();
        for (var j = rows - 1; j > 0; j--) {
          table.removeRow(j);
        }
        break;
      }
    }
  }

  if (!tableExists) {
    table = body.insertTable(0, [['Heading', 'Page Number']]); // Insert at the beginning of the body
    table.getRow(0).getCell(0).setText(tocTitle);
  }

  table.setBorderColor('#000000');
  table.setColumnWidth(0, 350);
  table.setColumnWidth(1, 100);

  var headings = body.getParagraphs();

  for (var i = 0; i < headings.length; i++) {
    var paragraph = headings[i];
    var textStyle = paragraph.getHeading();
    var text = paragraph.getText();

    if (textStyle === DocumentApp.ParagraphHeading.HEADING1) {
      var row = table.appendTableRow();
      var cell1 = row.appendTableCell(text);
      var cell2 = row.appendTableCell(String(table.getNumRows() - 1));

      cell1.setBackgroundColor('#D9EAD3');
      cell1.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      cell1.getChild(0).asText().setBold(true); // Set text in Heading 1 row to bold
      cell2.setBackgroundColor('#7E00FF');
      cell2.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    } else if (textStyle === DocumentApp.ParagraphHeading.HEADING2) {
      var row = table.appendTableRow();
      var cell1 = row.appendTableCell(text);
      var cell2 = row.appendTableCell(String(table.getNumRows() - 1));

      cell1.setBackgroundColor('#D9EAD3');
      cell1.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      cell1.setPaddingLeft(30); // Indent the text in Heading 2 row
      cell2.setBackgroundColor('#7E00FF');
      cell2.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    }
  }
}
