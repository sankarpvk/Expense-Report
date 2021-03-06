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
  var mailBody = "Hello Sankar, <br/><br/>Your expense till date for the month of " +currmonth+ " is " + sum+" Rs.";
  mailBody = mailBody+ "<br/>Please find the detailed report here <a href = \""+url+"\">Click here</a>";
  mailBody = mailBody+"<br/><br/>Regards,<br/>Google";
  //MailApp.sendEmail("sankar.potty@gmail.com","expense report", "your expense till date for the month of "+currmonth+ " is " + sum + " Rs.");
  MailApp.sendEmail("sankar.potty@gmail.com","Expense Report",mailBody,{'htmlBody':mailBody});
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

function DailyReport()
{
 var ret = SpendLimit(true);
}

function SpendLimit(dailyreport)
{
  var currmonth = getSheet();
  var sheet = SpreadsheetApp.getActive().getSheetByName(currmonth);
  var range = sheet.getRange("A2:C100");
  var data = range.getValues();
  
  var limitsheet = SpreadsheetApp.getActive().getSheetByName('Limits');
  var limit = limitsheet.getRange("A2:B10");
  var limitvalues = limit.getValues();
  //Logger.log(limitvalues[0][1]);
  
  var items = limitsheet.getRange("F1:P10");
  var itemlist = items.getValues();
  
   
  //Logger.log(limitvalues[1][1]);
  
  var fooddelivery =0;
  var utility=0;
  var grocery =0;
  var party=0;
  var eatout=0;
  var others=0;
  
  for (i in data)
  {
    var row = data[i];
   
   if (row[2]!='' && itemlist[0].indexOf(row[2])>=0)
    {
     fooddelivery = fooddelivery+row[1];
    }
   else if (row[2]!= '' && itemlist[1].indexOf(row[2])>=0)
    {
      utility=utility+row[1];
      
    }
        else if (row[2]!='' && itemlist[2].indexOf(row[2])>=0)
    {
     eatout = eatout+row[1];
    }
     else if (row[2]!='' && itemlist[3].indexOf(row[2])>=0)
    {
     grocery = grocery+row[1];
    }
     else if (row[2]!='' && itemlist[4].indexOf(row[2])>=0)
    {
     party = party+row[1];
    }
    else if (row[2]!='' && itemlist[5].indexOf(row[2])>=0)
    {
     others = others+row[1];
    }
  
  }
  
  if (dailyreport == true)
  {
    var reportresult = generateReport(fooddelivery,utility,eatout,grocery,party,others);
  }
  else if (fooddelivery >= limitvalues[0][1])
  {
   sendAlertNotification(limitvalues[0][0],limitvalues[0][1],fooddelivery);
  }
  if (utility >=  limitvalues[1][1])
  {
  sendAlertNotification(limitvalues[1][0],limitvalues[1][1],utility);
  }
  if (eatout >= limitvalues[2][1])
  {
  sendAlertNotification(limitvalues[2][0],limitvalues[2][1],eatout);
  }
  if (grocery >= limitvalues[3][1])
  {
  sendAlertNotification(limitvalues[3][0],limitvalues[3][1],grocery);
  }
  if (party >= limitvalues[4][1])
  {
  sendAlertNotification(limitvalues[4][0],limitvalues[4][1],party);
  }
  if (others >= limitvalues[5][1])
  {
  sendAlertNotification(limitvalues[5][0],limitvalues[5][1],others);
  }
  
  //sendAlertNotification(limitvalues[0][0],limitvalues[0][1],fooddelivery);
  //sendAlertNotification(limitvalues[1][0],limitvalues[1][1],utility);
}

function sendAlertNotification(type,limit,currentspends)
{

  var mailBody = "Hello Sankar, <br/><br/>You have crossed the current month expense limit for "+type;
  mailBody = mailBody+ "<br/><br/>Current Spends: "+currentspends;
  mailBody = mailBody+"<br/><br/>Monthly Limits: "+limit;
  mailBody = mailBody+"<br/><br/>Regards,<br/>Google";
  //MailApp.sendEmail("sankar.potty@gmail.com","expense report", "your expense till date for the month of "+currmonth+ " is " + sum + " Rs.");
  MailApp.sendEmail("sankar.potty@gmail.com","Alert! Expense threshold reached for "+type,mailBody,{'htmlBody':mailBody});
}

function generateReport(Fd,Uy,Eo,Gr,Py,Ot)
{
  var tableheader = getTableStyle();
  var total = Fd+Uy+Eo+Gr+Py+Ot;
  var mailBody = tableheader+ "<body>Hello Sankar, <br/><br/>Please find your daily expense report below";
  mailBody = mailBody+ "<br/><br/>";
  mailBody = mailBody + "<table><tr><th id=\"th01\"><b>Expense Type</b></th><th id=\"th01\"><b>Total Spend</b></th></tr>";
  mailBody = mailBody + "<tr id=\"tr01\"><td>Food Delivery</td><td>"+Fd+"</td></tr>";
  mailBody = mailBody + "<tr><td>Utility</td><td>"+Uy+"</td></tr>";
  mailBody = mailBody + "<tr id=\"tr01\"><td>Eat Out</td><td>"+Eo+"</td></tr>";
  mailBody = mailBody + "<tr><td>Grocery</td><td>"+Gr+"</td></tr>";
  mailBody = mailBody + "<tr id=\"tr01\"><td>Party</td><td>"+Py+"</td></tr>";
  mailBody = mailBody + "<tr><td>Others</td><td>"+Ot+"</td></tr>";
  mailBody = mailBody + "<tr id=\"tr01\"><td><b>Total</b></td><td><b>"+total+"</b></td></tr></table>"
  mailBody = mailBody+"<br/><br/>Regards,<br/>Google</body>";
 
 MailApp.sendEmail("sankar.potty@gmail.com","Daily Expense Report ",mailBody,{'htmlBody':mailBody});
}

function getTableStyle()
{
var TbStyle = "<head><style>table {font-family: arial, sans-serif;border-collapse: collapse;width: 50%;}"
TbStyle = TbStyle+"td, th {border: 2px solid #dddddd;text-align: left;padding: 8px} tr#tr01{background-color: #f1f1c1;} th#th01{background-color:#aedeff;}</style></head>";

return TbStyle;
}