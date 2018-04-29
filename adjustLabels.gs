// This converts Labels produced by Google Docs add on Avery Label Merge 
// https://chrome.google.com/webstore/detail/avery-label-merge/onacejilpbdgdhhoijnggpjolbdajepo?utm_source=permalink
// that are created for US Legal sheet sizes so they fit labels using A4 sheet sizes
// This specific code takes a Merged Doc formatted for Avery 5162 labels (14 per sheet)
// and converts it to fit onto Avery L7163 label sheet stock

function adjustLabels() {
  var body = DocumentApp.getActiveDocument().getBody();
  // set to A4 - credit  https://stackoverflow.com/a/45545118
  body.setPageHeight(841.89).setPageWidth(595.276).setMarginLeft(25);
  var tables = body.getTables();
  for each (var table in tables) {
    // enumerate rows, credit https://coderwall.com/p/k8iytg/
    numRows = table.getNumRows();
    for (var r = 0; r < numRows; r++) {
      // The label depth is 38.1mm but using Point eqivalent of 36.3mm because there's .7mm cell padding
      // hmmm, perhaps those figures don't quite add up yet!?!
      table.getRow(r).setMinimumHeight(103);
    }
  }
}

