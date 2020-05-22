function doGet() {
  var html = HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('Web App').setSandboxMode(HtmlService.SandboxMode.NATIVE);
    
  return html
}
