function callsdate() {
 var lvss = SpreadsheetApp.getActiveSpreadsheet();
 var lvSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Filter");

 var keyid = "----------------------";
 var key = "-----------------------";
 var url= 'https://api.paypal.com/v1/invoicing/search'; 
 var headers = {"Authorization" : "Basic " + Utilities.base64Encode(keyid + ':' + key)}; 
 var options = { "method": "POST", "headers": headers,"muteHttpExceptions": true, "contentType" : "application/json", "payload" : JSON.stringify({"status" : ["PAID"], "page" : 0, "page_size": 100})}; 
 var response = UrlFetchApp.fetch(encodeURI(url), options);
 var dataAll = JSON.parse(response.getContentText()); 
 //Logger.log(response);
var arr = [];
for (var i=0; i<100; i++){ 
var id = dataAll.invoices[i].id;
var url2= 'https://api.paypal.com/v1/invoicing/invoices/' + id; 
 
 var options2 = { "method": "GET", "headers": headers,"muteHttpExceptions": true, "contentType" : "application/json"}; 
 var responses = UrlFetchApp.fetch(encodeURI(url2), options2);
 
var dataAlls = JSON.parse(responses.getContentText()); 
Logger.log(dataAlls);
var date = dataAlls.payments[0].date;
var details = dataAlls.items[0].name;

var status = dataAll.invoices[i].status;
var amount = dataAll.invoices[i].total_amount.value;
var curr = dataAll.invoices[i].total_amount.currency;

lvSheet.appendRow([date,id,status,amount,curr, details]);  }
}
