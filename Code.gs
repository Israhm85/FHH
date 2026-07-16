function doGet(e) {
  var page = e && e.parameter && e.parameter.page;

  if (page === 'placard') {
    return HtmlService.createTemplateFromFile('placard')
      .evaluate()
      .setTitle('East Coast Placard Generator');
  }

  // Default: serve the email builder
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle("Gelson's Email Builder");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

