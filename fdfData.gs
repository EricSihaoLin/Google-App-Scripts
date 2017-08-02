/* FDF Data Extraction
 * By Eric Lin
 *
 * This script SHOULD live in a Google Spreadsheet's Script Editor
 * 
 * It will scan the user's emails for anything from ServiceLink and FDF related
 * And put it in the spreadsheet accordingly
 */

function extractData() {

  var restart = true;
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetHeader = sheet.getRange(1, 1, 1, 8);
  
  if(sheetHeader.isBlank())
  {
    sheetHeader.setValues([["FDF Number", "Date Submitted", "Department", "Requested Action", "Short Description", "Start Date", "Attachment Name", "Attachment URL"]]);
  }

  var labelList = GmailApp.getUserLabels();
  var labelExists = false;
  for (var i=0; i<labelList.length; i++)
  {
    if(labelList[i].getName() === "Processed FDFs")
    {
      labelExists = true;
    }
  }
  if (labelExists && restart)
  {
    GmailApp.getUserLabelByName("Processed FDFs").deleteLabel();
    restart = false;
    labelExists = false;
  }
  if (!labelExists)
  {
    GmailApp.createLabel("Processed FDFs");
  }
  
  var label = GmailApp.getUserLabelByName("Processed FDFs");
  var threads = GmailApp.search("!label:processed-fdfs from:(servicelink@nyu.edu) FDF");
  //var threads = GmailApp.search("!label:processed-fdfs from:(william.pride@nyu.edu) FDF");

  for (var i=0; i<threads.length; i++)
  {
    var messages = threads[i].getMessages();

    for (var j=0; j<messages.length; j++)
    {
      var from = messages[j].getFrom();
      if(messages[j].getFrom().match("servicelink@nyu.edu"))
      //if(messages[j].getFrom().match("william.pride@nyu.edu"))
      {
        var msg = messages[j].getBody();
        var sub = messages[j].getSubject();
        var dat = messages[j].getDate();
      
        var numFDF = extractFDFNumber(sub);
        var dept = extractDepartmentName(msg);
        var action = extractRequestedAction(msg);
        var desc = extractShortDescription(msg);
        var start = extractStartDate(msg);
        var attachName = extractAttachmentName(msg);
        var attachURL = extractAttachmentURL(msg);

        sheet.appendRow([numFDF, dat, dept, action, desc, start, attachName, attachURL]);
      }
    }
    threads[i].addLabel(label);
  }
}

function extractFDFNumber(subject) {
  return subject.match(/\d+/)[0];
}

function extractDepartmentName(body) {
  try{
    var name = body.match(/Department Name:\s*(.*?)<\/div>/)[1];
    return name;
  }
  catch(err)
  {
    return "";
  }
}

function extractRequestedAction(body) {
  try{
    var action = body.match(/Requested Action:\s*(.*?)<\/div>/)[1];
    return action;
  }
  catch(err)
  {
    return "";
  }
}

function extractShortDescription(body) {
  try{
    var desc = body.match(/Short Description:\s*(.*?)<\/div>/)[1];
    return desc;
  }
  catch(err)
  {
    return "";
  }
}

function extractStartDate(body) {
  try{
    var start = body.match(/Start Date:\s*(.*?)<\/div>/)[1];
    return start;
  }
  catch(err)
  {
    return "";
  }
}

function extractAttachmentName(body) {
  try{
    var name = body.match(/Attachment:\s*<a\s*href=".*?>(.*?)<\/a><\/div>/)[1];
    return name;
  }
  catch(err)
  {
    return "";
  }
}

function extractAttachmentURL(body) {
  try{
    var url = body.match(/Attachment:\s*<a\s*href="(.*?)"\s*.*?<\/div>/)[1];
    return url;
  }
  catch(err)
  {
    return "";
  }
}