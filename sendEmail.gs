

function sendEmail(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('Sheet1');
  
  var n=sheet1.getLastRow();
  for (var i=2; i<n+1; i++){
    var emailaddress = sheet1.getRange(i,1).getValue();
    var name = sheet1.getRange(i,2).getValue();
    var location = sheet1.getRange(i,3).getValue();
    var position = sheet1.getRange(i,4).getValue();
    var subject = "Hello, ";
    var message = "***********************."
    MailApp.sendEmail(emailaddress, subject, message);

    // sheet1.deleteRow(i);
    recharge(emailaddress, name, location, position, i);
  }

}

function recharge(emailaddress, name, location, position, i){
  var sss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet3 = sss.getSheetByName("Sheet3")
  var n = sheet3.getLastRow();
  sheet3.getRange(n+1,1).setValue(emailaddress);
  sheet3.getRange(n+1,2).setValue(name);
  sheet3.getRange(n+1,3).setValue(location);
  sheet3.getRange(n+1,5).setValue(position);
  var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")
  var endDate = date
  sheet3.getRange(n+1,4).setValue(endDate)
  if(i == 2){
    sheet3.getRange(n+1,1).setBackground("yellow");
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function onOpen() {
    const spreadsheet = SpreadsheetApp.getActive();
    let menuItems = [
        {name: 'Gather emails', functionName: 'gather'},
    ];
    spreadsheet.addMenu('Charity', menuItems);
}


function gather() {
    let messages = getGmail();
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let curSheet = ss.getSheetByName('Sheet2');

    messages.forEach(message => {curSheet.appendRow(parseEmail(message))});
}

function getGmail() {
    const query = 'label:inbox';

    let threads = GmailApp.search(query);

    let label = GmailApp.getUserLabelByName("done");
    if (!label) {label = GmailApp.createLabel("done")}

    let messages = [];

    threads.forEach(thread => {
        messages.push(thread.getMessages()[0].getPlainBody());
        label.addToThread(thread);
    });

    return messages;
}

function parseEmail(message){
    let parsed = message.replace(/,/g,'')
        .replace(/\n*.+:/g,',')
        .replace(/^,/,'')
        .replace(/\n/g,'')
        .split(',');

    let result = [0,1,2,3,4,6].map(index => parsed[index]);

    return result;
}




