// Trenzo orders -> Google Sheet row + e-mail alert
var OWNER_EMAIL = 'mohommed.adeeb20@gmail.com';

// Shows the website when someone opens the web app link
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Trenzo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  // Only record + e-mail when the customer pressed the green "Place your order on WhatsApp" button
  if (d.trigger !== 'whatsapp_button') return ContentService.createTextOutput('ignored');
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);   // makes sure two customers never get the same number
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Orders') ||
           SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Order ID','Date & Time','Name','Address','City','District','Phone','Second Phone',
                  'Items','Qty','Items Total','Delivery','Total','Payment','Status']);
    sh.setFrozenRows(1);
  }
  // Order number = highest number already in column A + 1 (1, 2, 3, 4 ...)
  var last = sh.getLastRow(), n = 0;
  if (last > 1) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) { var v = Number(ids[i][0]); if (v > n) n = v; }
  }
  d.orderId = n + 1;
  var when = Utilities.formatDate(new Date(d.time), 'Asia/Colombo', 'yyyy-MM-dd HH:mm:ss');
  sh.appendRow([d.orderId, when, d.name, d.address, d.city, d.district, "'" + d.phone,
                d.phone2 ? "'" + d.phone2 : '', d.items, d.quantity, d.itemsTotal, d.delivery,
                d.total, d.payment, 'New']);
  MailApp.sendEmail({
    to: OWNER_EMAIL,
    subject: 'New Trenzo order ' + d.orderId + ' - ' + d.name,
    body: 'You have a new order!\n\n' + d.message + '\n\nOrder ID: ' + d.orderId + '\nTime: ' + when
  });
  SpreadsheetApp.flush();
  lock.releaseLock();
  return ContentService.createTextOutput('ok');
}
