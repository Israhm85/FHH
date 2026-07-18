function doGet(e) {
  return HtmlService.createTemplateFromFile('placard')
    .evaluate()
    .setTitle('FHH Placard Generator');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

