function sendEmail(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();  //ss is linked spreadsheet
  var sheetForm = ss.getSheetByName('Form Responses 1');
  var rowForm = sheetForm.getLastRow();  //most recently filled row number
  var data = sheetForm.getDataRange().getValues();
  var name = data[rowForm-1][2];
  var email = data[rowForm-1][1];
  var phoneNum = data[rowForm-1][3];
  var p = sheetForm.getRange(rowForm,5);
  var subject="Thanks for registering";
  var htmlgood ="<body>"+ p.getValue()+
  "<br>Dear "+ name+",</br>"+
  "<br>Tokens are available! You may leave for the hostel in next 30 minutes</br>"+
  "<br>Have a nice day!"+
  "<br><br>Best Wishes,"+
  "<br>Security @ IIT Bhilai"+
  "</body>";

  var htmlbad ="<body>"+
  "<br> Tokens are in use! Due to social distancing norms we are allowing outings to only 5 students at once. We will notify you your slot soon. </br>"+
  "<br><br>Best Wishes,</br></br>"+
  "<br>Security @ IIT Bhilai"+
  "</body>";

  if (rowForm==2){
    p.setValue(5);
  }
  else{
    if (rowForm<=6){
    p.setValue(sheetForm.getRange(rowForm-1,5).getValue()-1);
    }
    else{ var i=0;
      
        while(i<data.length && sheetForm.getRange(i+1,9).getValue()!='Process Complete'){
          i++;
        }
        if (i!=data.length){
          p.setValue(sheetForm.getRange(i+1,5).getValue());
          sheetForm.getRange(i+1,9).setValue('Token Reused');
          sheetForm.getRange(i+1,5).setBackgroundRGB(255,0,0);
    }}
    }
    if(p.getValue()==''){
      p.setValue(0);
    }

  
  if (p.getValue()>0){
      sheetForm.getRange(rowForm,7).setValue("Notified");
      MailApp.sendEmail({to: email, subject: subject, htmlBody: htmlgood});
  }
  if (p.getValue()==0){
      sheetForm.getRange(rowForm,7).setValue("Will be notified");
      MailApp.sendEmail({to: email, subject: subject, htmlBody: htmlbad});
  }
}

function doGet(e) {
var htmlOutput = HtmlService.createTemplateFromFile("button.html");
htmlOutput.url = ScriptApp.getService().getUrl();
var ss = SpreadsheetApp.getActiveSpreadsheet();  //ss is linked spreadsheet
var sheetForm = ss.getSheetByName('Form Responses 1');
var rowForm = sheetForm.getLastRow();
var data = sheetForm.getDataRange().getValues();
  var name = data[rowForm-1][2];
  var subject="Thanks for registering";
  var htmlgood ="<body>"+
  "<br> Dear "+ name+",</br>"+
  "<br>Tokens are available! You may leave for the hostel in next 30 minutes</br>"+
  "<br>Have a nice day!"+
  "<br><br>Best Wishes,"+
  "<br>Security @ IIT Bhilai"+
  "</body>";
for(i=0; i<data.length; i++){

if(e.parameter['phonein']==data[i][3] && sheetForm.getRange(i+1,7).getValue()=="Notified" && sheetForm.getRange(i+1,8).getValue()=="" && 
sheetForm.getRange(i+1,9).getValue()==""){
  sheetForm.getRange(i+1, 8).setValue("Left");
  break;
}

if(e.parameter['phoneout']==data[i][3] && sheetForm.getRange(i+1,9).getValue()=="" ){
  sheetForm.getRange(i+1,9).setValue("Process Complete");
  var j=0;
  while(j<data.length && sheetForm.getRange(j+1,7).getValue()!="Will be notified"){
    j++;
  }

  if(j!=data.length){
  sheetForm.getRange(j+1,7).setValue("Notified");
  var k=0;
  while(k<data.length && sheetForm.getRange(k+1,9).getValue()!="Process Complete"){
    k++;
  }
  if(k!=data.length){
    var token = sheetForm.getRange(k+1,5).getValue();
    sheetForm.getRange(k+1,9).setValue("Token Reused");
    sheetForm.getRange(k+1,5).setBackgroundRGB(255,0,0);
    sheetForm.getRange(j+1,5).setValue(token);
    var email=sheetForm.getRange(j+1,2).getValue();
    MailApp.sendEmail({to: email, subject: subject, htmlBody: htmlgood});
  }
  }
  
  break;
  }
}

return htmlOutput.evaluate();

}
