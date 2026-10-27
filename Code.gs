function doGet(e) {
  return HtmlService.createTemplateFromFile('placard')
    .evaluate()
    .setTitle('FHH Placard Generator');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getEastCoastCatalogJson_() {
  var products;
  try {
    products = getEastCoastProducts_();
  } catch (err) {
    // Invalid EAST_COAST_PRODUCTS must not break the Manual form; the error is
    // still reported when sheet orders are loaded.
    products = getEastCoastProducts_(true);
  }
  return JSON.stringify(products).replace(/</g, '\\u003c');
}
