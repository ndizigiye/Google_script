/**
 *
 *
 *
 * Armand Ndizigiye
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++){
    var row = values[i];
    Logger.log(row);
  }
};

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Read Data",
    functionName : "readRows"
  }];
  var entries2 = [{
    name : "Import Agenda",
    functionName : "run"
  }];
  sheet.addMenu("Script Center Menu", entries);
  sheet.addMenu("Import Agenda", entries2);

};

/**
 * Searchs in the sheet and return  an array of 
 *the matching dates + column number
 *@return array of the matching dates + column number
 */
function getColumn(){

var sheet = SpreadsheetApp.getActiveSheet();
var sheetDates = getDates();
var agendaStartDates  = new Array();
var agendaStartDatesString  = new Array();
var events = getAllEvents();
var lookup = {};
var matchingDates = {};
  
  for (var i in events) {
  agendaStartDates[i] = events[i][0].setHours(0,0,0,0);
  agendaStartDatesString[i] = new Date(agendaStartDates[i]).toString();
  }

 for (var j in agendaStartDates) {
      lookup[agendaStartDatesString[j]] = events[j][1];
  }
  var j = 0;
  for (var i = 0  in sheetDates) {
      if (typeof lookup[sheetDates[i]] != 'undefined') {
        matchingDates[sheetDates[i]] = parseInt(i)+1;
         j++;
          }
  }
  
  return matchingDates;
}

/**
 * Main function that prints the events (event title) on the spreadsheet
 * @return void
 */
function run(){
var row = Browser.inputBox("row number:");
var events =  getAllEvents();
var matchingDates = getColumn();
var columnArray = {};
var sheet = SpreadsheetApp.getActiveSheet();
  
    for(var i = 0;i<events.length;i++){
      var row2 = row;
      for(var j in columnArray ){
        if (j.split(" ")[0] == (matchingDates[events[i][0]]+1)){
          row2 ++;
          }
       }
        var ColumnNumber = matchingDates[events[i][0]]+1;
        var eventTitle = events[i][1];
        var range = sheet.getRange(row2,ColumnNumber,1,1);
        range.setValue(eventTitle).setBackgroundColor("#00FF00");
       columnArray[(ColumnNumber+" "+i)] = "set";
    }
}

/**
 * Get all the dates in the current sheet. 
 * @return array of strings of dates
 */
function getDates() {
 var sheet = SpreadsheetApp.getActiveSheet();
 var range=sheet.getRange(2,2,1,400000);
 var sheetDatesArray = new Array();
 sheetDatesArray =  range.getValues().toString().split(",");
 return sheetDatesArray;
}

/**
 * Get the date and title of all events.The events with more than one day are supported. 
 * @return array of event day + event title
 */
function getAllEvents() {
  var joined = new Array();
  var newArray = new Array();
  var existingArray = new Array();
  var events = getEventsFeeds();
  var k = 0;
  for (var i in events){
    var eventTitle = events[i][0];
    var eventStartTime = events[i][1];
    var eventEndTime  = events[i][2];
    existingArray[i] = [eventStartTime,eventTitle];
    var days = Math.round(((eventEndTime-eventStartTime)+0.01)/86400000);
    var arrayLength = events.length;
    if(days >= 1 ){
      for (var j = 0; j < days; j++){
        var newStartDay =  eventStartTime.getTime()+ (86400000*(j+1));
        newArray[k] = [new Date(newStartDay),eventTitle];
         k++;
      }
     
    }
  }
  
  joined = existingArray.concat(newArray);
  
  return joined;
}

/**
 * Retrieves all calender events from a given calendar private adress
 * These function returns then a array containing the event title, event startdate and event end date.
 * @return array
 */
function getEventsFeeds(){

  var doc = UrlFetchApp.fetch("https://www.google.com/calendar/feeds/ts9e3cufie0q39p7k28q9a63k8@group.calendar.google.com/private-daf165ea84695a4c0cfcdd17ef3bfaa4/basic").getContentText();
  var xml = Xml.parse(doc);
  var feed = xml.feed;
  var entries = feed.getElements("entry");
  var events = new Array();
  
  for(var i in entries){
    var summary = entries[i].getElement("summary").getText();
    var eventStartTime = "";
    var eventEndTime = "";
    var title = entries[i].getElement("title").getText();
    if(summary.indexOf("&nbsp") > -1){
      eventStartTime = new Date(summary.split("&nbsp")[0].split(" to")[0].split("When: ")[1]);
     eventEndTime = new Date(summary.split("&nbsp")[0].split("to ")[1]);
    }
    else if(summary.indexOf("am&nbsp") > -1 || summary.indexOf("pm&nbsp")){
      eventStartTime = new Date(summary.split(" ",5)[0].split("When: ")[1]);
      eventEndTime = new Date (summary.split(" ",5)[0].split("When: ")[1]);
    }
    else{
      eventStartTime = new Date(summary.split("<br>")[0].split("When: ")[1]);
      eventEndTime = new Date (summary.split("<br>")[0].split("When: ")[1]);
    }
    
    events[i] = [title,eventStartTime,eventEndTime];
      }

  return events;
    
};


