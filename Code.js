function sendPaypalInvoice(e) {
 var invoicesent=0;
  var errorMessage='No Error Found';
  var client_id='---------------------';
  var secret_id='---------------------';

  try{
    var authorizationToken;
    var authorizationObj = getAuthorizationToken(client_id,secret_id);
    if(authorizationObj.error==true){
      errorMessage=authorizationObj.message;
      throw(new Error(errorMessage));
    }
    authorizationToken=authorizationObj.access_token;
 var lvss = SpreadsheetApp.getActiveSpreadsheet();
 var lvSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form");
  // get ALL the data from this sheet
  var lvdata = lvSheet.getDataRange().getValues();
  // check how many rows of data
  var lvLastRow = lvSheet.getLastRow();
for (var i=1; i<lvLastRow; i++){
      var tmail = e.values[1];
      var sname = e.values[2];
      var pname = e.values[3];
      var pmail = e.values[5];
      var pphone = e.values[4];
      var pro = e.values[13];
      var amt = e.values[8];
      var curr = e.values[9];
      var tenure = e.values[10];
      var cellrange = lvSheet.getRange(i+1, 19);
      var cellrangetwo = lvSheet.getRange(i+1, 20);
 
}
  var data={};
  data.logo_url="https://d138zd1ktt9iqe.cloudfront.net/-----------";
  //data.cc_info=[{email:'support@test.com'}];
  data.merchant_info={};
  data.merchant_info.email="abc@test.com";
  data.merchant_info.business_name="test Pvt. Ltd.";
  data.merchant_info.website="www.test.com" ;
 
  var tempBillingInfo={};  
  tempBillingInfo.first_name=pname;
  
  var temp={};
  data.items=[];
  
  temp.name=sname;
    
  temp.quantity= 1
  temp.unit_price={
        "currency": curr,
        "value": amt
      }
  data.items.push(temp);   
    

  var error='';
  var errorFlag=0;
  if(tempBillingInfo&&tempBillingInfo.email){
    data.billing_info=[];
    data.billing_info.push(tempBillingInfo);
}
  
 
      if(data.error){
        throw(new Error(data.message));
      }
      //check and validate invoiceData:= check valid email and the invoice amount
    
      var draftInvoiceObject=createInvoiceDraft(data,authorizationToken);
      if(draftInvoiceObject.error){
        
        throw(new Error(JSON.stringify(draftInvoiceObject)));
        
      }
      var invoiceConfirmation = sendDraftInvoice(draftInvoiceObject.id,authorizationToken);
      if(invoiceConfirmation.error){
   
        throw(new Error(JSON.stringify(invoiceConfirmation)));

      }
    invoicesent=1;
    cellrange.setValue("Invoice Sent")
    cellrangetwo.setValue(draftInvoiceObject.id);
    
  var subject = "Payment link - Test Pvt. Ltd: " + sname;
  var cc = pmail;
  var htmlBody1 = "";
  htmlBody1 += '<a href="https://www.test.com/">' + '<img src="----------" border="0">' + '</a>' + '<br />';
  htmlBody1 += '<p>' + 'Dear ' + pname + ',' + '</p>';
  htmlBody1 += '<p>' + 'https://www.paypal.com/invoice/p/#' + draftInvoiceObject.id + '</p>';
  htmlBody1 += '<p>' + 'You can click on the link and make the fee payment through debit card/credit card/PayPal credits.' + '</p>';
  htmlBody1 += '<br>' + '<b>' + 'Details :' + '</b>' ;
  htmlBody1 += '<br>' + 'Product - ' + pro;
  htmlBody1 += '<br>' + 'Tenure - ' + tenure;
  htmlBody1 += '<p>' + 'Please do let me know in case of any concerns or any difficulty which you face' + '</p>';
  htmlBody1 += '<br>' + 'Best wishes'
  htmlBody1 += '<br>' + 'Team Business' + '<br>' + '<a href="http://www.test.com/">' + 'www.test.com' + '</a>';
    MailApp.sendEmail(pmail, subject,"" , {htmlBody : htmlBody1 , 'cc': tmail + ',support@test.com'});
   
    
  }catch(err){
    cellrange.setValue(err.message);
    
      Logger.log(err.message);
      MailApp.sendEmail("shruti.asthana@test.com","Contact Admin- Automated invoicing Api not working",err.message);
    
    }
  
  //range.setValues(values);
  
}
function sendDraftInvoice(invoiceId,authorizationToken){
  head = {
    'Authorization':"Bearer "+ authorizationToken,
    'Content-Type': 'application/json'
  }
  
   params = {
    headers:  head,
    method : "post",
    muteHttpExceptions: true
  }
  tokenEndpoint='https://api.paypal.com/v1/invoicing/invoices/'+invoiceId+'/send';
  request = UrlFetchApp.getRequest(tokenEndpoint, params); 
  response = UrlFetchApp.fetch(tokenEndpoint, params); 
  
  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();
   var invoiceResponse={};
  if (responseCode === 202) {
    
    invoiceResponse.error=false;
    return invoiceResponse;

    } else {
      invoiceResponse.error=true;
      invoiceResponse.message=Utilities.formatString("Request failed. Expected 202, got %d: %s", responseCode, responseBody);
      return invoiceResponse;
    }
  
  
}


function createInvoiceDraft(invoiceData, authorizationToken){
  head = {
    'Authorization':"Bearer "+ authorizationToken,
    'Content-Type': 'application/json'
  }
  params = {
    headers:  head,
    method : "post",
    muteHttpExceptions: true,
    payload:JSON.stringify(invoiceData)    
  }
  tokenEndpoint='https://api.paypal.com/v1/invoicing/invoices';
  request = UrlFetchApp.getRequest(tokenEndpoint, params); 
  response = UrlFetchApp.fetch(tokenEndpoint, params); 
  
  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();

  var invoiceResponse={};
  if (responseCode === 201) {
    var responseJson = JSON.parse(responseBody);
    invoiceResponse.error=false;
    invoiceResponse.id=responseJson.id;
    } else {
      invoiceResponse.error=true;
      invoiceResponse.message=Utilities.formatString("Request failed. Expected 200, got %d: %s", responseCode, responseBody);
    }
  
  return invoiceResponse;
}

function getAuthorizationToken(client_id,secret_id){
  var tokenEndpoint = "https://api.paypal.com/v1/oauth2/token";
    var head = {
      'Authorization':"Basic "+ Utilities.base64Encode(client_id+':'+secret_id),
      'Accept': 'application/json',
      'Content-Type': 'application/x-www-form-urlencoded'
    }
    var postPayload = {
        "grant_type" : "client_credentials"
    }
    var params = {
        headers:  head,
        contentType: 'application/x-www-form-urlencoded',
        method : "post",
        muteHttpExceptions: true,
        payload : postPayload      
    }
    var request = UrlFetchApp.getRequest(tokenEndpoint, params); 
    var response = UrlFetchApp.fetch(tokenEndpoint, params); 
    var responseCode = response.getResponseCode()
    var responseBody = response.getContentText()

    if (responseCode === 200) {
      var tokenResponse={};
      var responseJson = JSON.parse(responseBody);
      if(responseJson&&responseJson.error){
        tokenResponse.error=true;
        tokenResponse.message=responseJson.error;
        return tokenResponse;
      }
      if(responseJson.access_token){
        tokenResponse.error=false;
        tokenResponse.access_token=responseJson.access_token;
        return tokenResponse;
      }
     tokenResponse.error=true;
     tokenResponse.message='Access Token not found';
     return tokenResponse;
    } else {
        var tokenResponse={};
        tokenResponse.error=true;
        tokenResponse.message=Utilities.formatString("Request failed. Expected 200, got %d: %s", responseCode, responseBody);
        return tokenResponse;
      }
}

