//Main function to calculate the expense till date for the current month and to send it as a report in an email. The email is currently scheduled to run every sunday 10:30 AM
function myFunction() {
  var currmonth = getSheet();
  var sheet = SpreadsheetApp.getActive().getSheetByName(currmonth);
  var range = sheet.getRange("A2:C100");
  var data = range.getValues();
  var sum =0;
  for (i in data)
  {
    var row = data[i];
    sum = sum+row[1];
  }
  var url = SpreadsheetApp.getActive().getUrl();
  //Logger.log(url);
  var mailBody = "Hello Sankar, Your expense till date for the month of " +currmonth+ " is " + sum+" Rs.>";
  mailBody = mailBody+ "<br/>Please find the detailed report here <a href = \""+url+"\">Click here</a>";
  //MailApp.sendEmail("sankar.potty@gmail.com","expense report", "your expense till date for the month of "+currmonth+ " is " + sum + " Rs.");
  MailApp.sendEmail("sankar.potty@gmail.com","expense report",mailBody,{'htmlBody':mailBody});
}

//Function to determine the current month and to select the sheet based on that.
function getSheet()
{
 var now = new Date()
 var month = now.getMonth();
 var activesheet = '';
 if (month == 0)
 {
   activesheet = 'Jan';
 }
 else if (month == 1)
 {
   activesheet = 'Feb';
 }
 else if (month == 2)
 {
 activesheet = 'Mar';
 }
  else if (month == 3)
 {
 activesheet = 'Apr';
 }
  else if (month == 4)
 {
 activesheet = 'May';
 }
  else if (month == 5)
 {
 activesheet = 'Jun';
 }
  else if (month == 6)
 {
 activesheet = 'Jul';
 }
  else if (month == 7)
 {
 activesheet = 'Aug';
 }
  else if (month == 8)
 {
 activesheet = 'Sep';
 }
  else if (month == 9)
 {
 activesheet = 'Oct';
 }
  else if (month == 10)
 {
 activesheet = 'Nov';
 }
  else
 {
 activesheet = 'Dec';
 }
 

 //Logger.log("current month is" + month);
 return activesheet;
}
