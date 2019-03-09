//  Setup Instructions (continued):
//
//  For Google Sheets to receive data from the Tilt app
//  deploy script as web app from the "Publish" menu and set permissions. Note that you are now the owner and "developer" of the app.
//
//  1) Got to "Publish" menu and select "Deploy as web app..."
//   
//  2) In the dialog box, set "Who has access to the app:" to "Anyone, even anonymous" and click "Deploy".
//
//  3) A dialog box will appear. Select "Review Permissions". Another dialog box will appear. Select your Google Account.
//
//  4) A dialog box with "This app isn't verified" will appear. Select "Advanced" then select "Go to Tilt Cloud Template for Tilt App 1.6+ (unsafe)"
//
//  5) A dialog box with permission requests will appear. Select "Allow".
//
//  6) A dialog box confirming the app has been published will appear. Note: Do NOT use the cloud URL shown in the dialog, see next step.
//
//  7) Close Google Scripts tab and return to Google Sheets. Use the new "Tilt" menu to email yourself the cloud URL.

var SHEET_NAME = "Data";
var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

// If you don't want to expose either GET or POST methods you can comment out the appropriate function
function doGet(e){
  return ContentService
  .createTextOutput("Enter the following link into the Cloud URL settings (under the gear icon in iOS/Android app):" + ScriptApp.getService().getUrl())
          .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e){
  return handleResponse(e);
}

//used for testing without a Tilt
function testBeer(){
  var e = {
  "parameter": {
  "Beer": "some regular beer",
  "Temp": 65,
  "SG":1.050,
  "Color":"BLUE",
  "Comment":"noahbaron@gmail.com",
  "Timepoint":42728.4267217361
  }
  };
  handleResponse(e);
}

function handleResponse(e) {
  //prevent simultaneous writes
  var lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    // next set where we write the data - you could write to multiple/alternate destinations 
    var masterDoc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    var beersSheet = masterDoc.getSheetByName("Beers");
    var nextBeerRow = (beersSheet.getLastRow()+1).toFixed(0);
    //app expects beer name to be followed by a comma and beer ID (except for new beers)
    var beerName = e.parameter.Beer.split(",");
    //if beer name is blank, give beer a default name
    if (beerName[0] == "") {
      beerName[0] = "Untitled";
    }
    var tiltColor= e.parameter.Color;
    var comment = e.parameter.Comment;
    var beerIds = beersSheet.getRange("A:C").getValues();
    var beerId = null;
    var email = null;
    var doclongURL = "";
    //get Sheets ID if it exists
    for (var i = 0; i < beerIds.length; i++) {
        if (e.parameter.Beer.toLowerCase() == beerIds[i][0].toLowerCase()) {
            beerId = beerIds[i][1];
        }
    }
    
    var doc = null;
    //check if this is a new beer or existing beer
    if(beerId == null){
      //check if comment field has an @ symbol for an email address
      if (comment.indexOf("@") > -1){
      var settingsSheet = masterDoc.getSheetByName("Settings");
      var sheetTemplate = settingsSheet.getRange("B1").getValue();
      var driveTemplate = DriveApp.getFileById(sheetTemplate); //file ID of template
      var driveDoc = driveTemplate.makeCopy(beerName[0] + " (" + tiltColor + " TILT)");
      driveDoc.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      doc = SpreadsheetApp.open(driveDoc);
      beerId = doc.getId();
      //add beer to 'Beers' tab
      doclongURL = doc.getUrl();
      beerName[1] = nextBeerRow;
      beersSheet.appendRow(["",beerId,doc.getUrl()]);
      beersSheet.getRange("A" + nextBeerRow).setNumberFormat('@STRING@');
      beersSheet.getRange("A" + nextBeerRow).setValue(beerName.join());
      driveDoc.addEditor(comment);
      MailApp.sendEmail({to : comment,
                         replyTo : "info@baronbrew.com",
                         subject : "Tilt™ Hydrometer Log for " + beerName[0],
                         body : 'View and edit your data here with Google Sheets: ' + doclongURL + ". To log data to the same sheet using another Tilt Pi, enter the following name as your beer name: " + beerName.toString() + " (Be sure to include comma and number afterward.)",
                         name : "Tilt Customer Service",
                        });
                           
      e.parameter.Comment = "";
      }
      else{
       //advise user to enter email into comment field
      return ContentService
      .createTextOutput(JSON.stringify({result:"</br>" + beerName[0] + "</br><strong>TILT | " + tiltColor + "</strong></br>Enter your GMAIL email address as a comment below to start a new brew log.", beername:beerName.toString(), tiltcolor:tiltColor}))
          .setMimeType(ContentService.MimeType.JSON);
    }
  }
    else{
      doc = SpreadsheetApp.openById(beerId);
      var editors = doc.getEditors();
      doclongURL = doc.getUrl();
      //check if comment field has a web address prefix used to transmit link to raspberry pi
      if (comment.indexOf("http") > -1){
        MailApp.sendEmail({to : editors[1].getEmail(),
                         replyTo : "info@baronbrew.com",
                         subject : "Tilt™ Pi Setup Complete",
                         body : "You may now access your Tilt Pi from your local WiFi network at http://tiltpi.local:1880/ui or at the following address: " + comment,
                         name : "Tilt Customer Service",
                        });
        e.parameter.Comment = "";
      }
    }
 
    var sheet = doc.getSheetByName("Data");    
    e.parameter.Beer = beerName[0]; //remove beer name unique identifier when posting to sheet
    // we'll assume header is in row 1 but you can override with header_row in GET/POST data
    var headRow = e.parameter.header_row || 1;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow()+1; // get next row
    var row = []; 
    // loop through the header columns
    for (i in headers){
      if (headers[i] == "Timestamp"){ // special case if you include a 'Timestamp' column
        row.push(new Date());
      } else { // else use header name to get data
        row.push(e.parameter[headers[i]]);
      }
    }
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    // return success results
    return ContentService
    .createTextOutput(JSON.stringify({result:"</br>" + beerName.toString() + "</br><strong>TILT | " + tiltColor + "</strong></br>Success logging data to the cloud. (row: " + nextRow + ")", beername:beerName.toString(), tiltcolor:tiltColor, doclongurl:doclongURL}))
          .setMimeType(ContentService.MimeType.JSON);
  } catch(e){
    // if error return this
    return ContentService
          .createTextOutput(JSON.stringify({result:"error", error: e}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
    lock.releaseLock();
  }
}



function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tilt')
    .addItem('View Cloud URL', 'menuItemURL')
    .addItem('Email Cloud URL', 'menuItemEmailURL')
    .addToUi();
  if(SCRIPT_PROP.getProperty("url")==null){
    setup();
  }
}

function setup() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  SCRIPT_PROP.setProperty("key", doc.getId());
  
    var html = HtmlService.createHtmlOutputFromFile('setup')
      .setTitle('Cloud Setup Instructions')
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
}

function menuItemURL() {
 
  if(ScriptApp.getService().getUrl()!=null){
    SCRIPT_PROP.setProperty("url", ScriptApp.getService().getUrl());
     SpreadsheetApp.getUi()
      .alert("Copy/Paste the following URL into the Cloud URL field in the Tilt app settings: " + ScriptApp.getService().getUrl());
  }
  else{
    SpreadsheetApp.getUi()
      .alert("Follow setup instructions in sidebar to deploy as web app");
  }
  
}

function menuItemEmailURL(){
  if(ScriptApp.getService().getUrl()!=null){
    SCRIPT_PROP.setProperty("url", ScriptApp.getService().getUrl());  
    MailApp.sendEmail(Session.getActiveUser().getEmail(), 'Tilt Cloud URL', "Copy/Paste the following URL into the Cloud URL field in the Tilt app settings: " + ScriptApp.getService().getUrl());
    SpreadsheetApp.getUi()
      .alert("Email sent to: " + Session.getActiveUser().getEmail());
  }
  else{
    SpreadsheetApp.getUi()
      .alert("Follow setup instructions in sidebar to deploy as web app");
  }
}