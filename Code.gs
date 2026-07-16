function doGet(e) {
  return HtmlService.createTemplateFromFile('placard')
    .evaluate()
    .setTitle('East Coast Placard Generator');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

