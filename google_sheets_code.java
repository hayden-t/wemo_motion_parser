//
// Analyse Belkin WeMo Motion Spreadsheet Logs for inactivity between 12am-Xam/pm (e.g. did not get up) and emails notifications.
// Code by www.httech.com.au
//
// This code must be added to a blank sheet via Tools->Script Editor...
// It needs to be kept seperately from the motion data spreadsheet which rolls over from time to time
// So this script gets the latest one last updated by chosen name
//
// Requires library "Moment" (Resources -> Libraries...->Find: MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48 )
//
// To Schedule: Resources -> Current Project's Triggers -> Add function "checkMotion" hourly
//
// Run once, "testEmail" function, to present dialog's to allow permissions for the script
//

//check for motion by chosen time in morning since midnight
var checkHour = 8; //check at nth hour, 8am for lack of motion since midnight

var sheetName = "New motion after quiet period";
var emailSubject = "Granny's Motion Sensor Notification";
var emailMessage = "Likely this is a false alarm, a problem with the internet, the sensor, or this script.\n\nNo Motion This Morning.\n\nLast: ";

var emails = ["email1@domain.com", "email2@domain.com","email3@domain.com"];//edit to suit needs


function getLatestSheet() {
 
  var files = DriveApp.searchFiles('title contains "'+sheetName+'"');
  
  if(files.hasNext()) {
    var file = files.next(); 

    Logger.log("Sheet Found Latest Updated: " + file.getLastUpdated());
    
    var spreadsheet = SpreadsheetApp.open(file);
    
    return spreadsheet;
 }
   Logger.log("No Sheet Found, Aborting");
   return false
}


//check for motion by chosen time in morning

var timeFormat = "MMM DD, YYYY at hh:mmA";

function sheetProcessor(test){  

  var moment = Moment.load();
  
  var mail = false;
  
  Logger.log("Checker Starting...");
  
  
  var ss = getLatestSheet();
  if(!ss)return;
  
  var sheet = ss.getSheets()[0];
  
  var lastRow = sheet.getRange(sheet.getLastRow(), 1).getValue();
  
  //Logger.log("Last Row = "+ lastRow);
  
  var now = moment();
  
  Logger.log("Now = " + now.format(timeFormat));
  
  
  var last = moment(lastRow, timeFormat);
  
  Logger.log("Last = "+ last.format(timeFormat));
 
  var checkAt = moment().hour(checkHour).minute(0).second(0);
  
  Logger.log("Check = " + checkAt.format(timeFormat));

  
  if(now > checkAt){//past check time

    var limit = moment().hour(0).minute(0).second(0);//midnight this morning
  
    Logger.log("Limit = " + limit.format(timeFormat));
    
    Logger.log("Time to Check...")
    
    if(last < limit){//no motion today
      
      Logger.log("No Motion Today");
      mail = true;
      
    }else Logger.log("Motion Today");
    
    if(test || mail){
      
      if(test)emailSubject = "Test: "+emailSubject;
      //test mail
      Logger.log("Sending Mail...");
       
     for (var i = 0; i < emails.length; ++i) {

       MailApp.sendEmail(emails[i], emailSubject, emailMessage + lastRow);
     }

    }
    
  }


}

function testEmail(){
  sheetProcessor(true);  
}
function checkMotion(){
  sheetProcessor(false);  
}
/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if(!spreadsheet)return;
  var entries = [{
    name : "Test Check Motion & Mail",
    functionName : "testEmail"
  }];
  spreadsheet.addMenu("Script Center Menu", entries);
};