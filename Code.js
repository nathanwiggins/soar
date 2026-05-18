function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SOAR - Project Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://drive.google.com/uc?id=1hGM6qVNuhjNcFbuB4eWyOwYbKDIOO66Y&export=download&format=png');
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
