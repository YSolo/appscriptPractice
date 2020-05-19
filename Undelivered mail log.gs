function logMailToTable() {
  const message = GmailApp.getInboxThreads(0, 1)[0].getMessages()[0];
  
  const body = message.getBody();
  let index = body.indexOf('Дата: ');
  let string = body.substring(index, index + 5);
  
  Logger.log(string);
}
