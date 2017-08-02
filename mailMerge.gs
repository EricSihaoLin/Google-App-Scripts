/* Mail Merge Script
 * By Eric Lin
 *
 * This script SHOULD live in a Google Spreadsheet's Script Editor
 * To access the script editor: Tools -> Script Editor
 * and then copy and paste this entire document into the code editor. Press save, then close out the script editor.
 * Close and reopen the document, the document should have all the features installed upon relaunch
 * 
 * This script will transform any spreadsheet into a mail merge
 * If this is your first time running the script, please run Mail Merge -> Setup Configuration
 */

//On open hook, runs every single time the spreadsheet is open
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Mail Merge')
      .addItem('Process Mail Merge', 'alertMerge')
      .addItem('Setup Configuration', 'alertConfig')
      .addToUi();
}

//Alert merge, confirmation dialogue to start Mail Merge process
function alertMerge() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.alert(
     'Please confirm',
     'Have you made sure that all the configuration is set correctly?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert(checkConfiguration());
  }
}

//Alert config, confirmation dialogue to start generating a Mail Merge Configuration sheet
function alertConfig() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to insert a configuration page into this Spreadsheet?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert(insertConfiguration());
  }
}

//Insert configuration sheet after making sure the sheet does not already exist
function insertConfiguration() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var conf = ss.getSheetByName("Mail Merge Configuration");
  if(conf != null){
    return "It seems like this Spreadsheet already has a sheet called Mail Merge Configuration, if you want to insert a new one, you must delete/rename the old one.";
  }
  
  conf = ss.insertSheet();
  conf.setName("Mail Merge Configuration");
  var confData = conf.getRange("A1:C16");
  var confPresets = [];
  //Some default data, don't worry about these.
  confPresets.push(["Configurations:", "ANYTHING IN YELLOW IS FOR YOU TO EDIT.", "Instructions:"]);
  confPresets.push(["Mail Merge Active Sheet Name:", "Test Merge", "Must be a sheet name from the bottom exactly. Otherwise will yell at you."]);
  confPresets.push(["Mail Merge Email Column (Column Letter):", "F", "Must be the column letter that stores all the emails addresses. Will validate every email address."]);
  confPresets.push(["Mail Merge # of Data Columns:", "6", "The number of data columns you have. If you have A-F columns filled, you have 6 columns."]);
  confPresets.push(["Mail Merge Send Mail From Name:", "Eric Lin", "The name you want to send as. You can use \"Blair Simmons\" or \"LaGuardia Co-op\""]);
  confPresets.push(["Mail Merge Send Mail From Email:", "eric.lin@nyu.edu", "By default, will be sent from NetID@nyu.edu. That's ugly, so use one of your aliases such as first.lastname@nyu.edu"]);
  confPresets.push(["Mail Merge Reply To:", "eric.lin@nyu.edu", "If you want clients to reply to student.tech.centers@nyu.edu instead of just you, you can do that here."]);
  confPresets.push(["Mail Merge Error/Success Log Email:", "eric.lin@nyu.edu", "In case if you have problems with this automation, I would require you to forward me the log so I can see what's wrong"]);
  confPresets.push(["Mail Merge BCC on Outgoing Email (Yes/No):", "Yes", "If you would like to be BCC'd on all the email this script sends, say Yes."]);
  confPresets.push(["Mail Merge BCC to Emails (Separated by Comma):", "eric.lin@nyu.edu", "Who would you like to BCC on these \"automated emails\"?"]);
  confPresets.push(["Mail Merge Use Pretty HTML (Yes/No):", "Yes", "If you want to use the pretty HTML interface with your email template"]);
  confPresets.push(["Mail Merge Subject Line (No Personalization):", "Extend Your Classroom to the LaGuardia Co-op!", "Customize the subject line for the email"]);
  confPresets.push(["Mail Merge Top Header Image URL (Ideal height is 50 pixels):", "http://wp.nyu.edu/eric/wp-content/uploads/sites/3202/2017/07/Header-2.png", "Customize the header image for the email. Use Pretty HTML must be Yes for this to work. Has to be a link to a picture that ends in .jpg, .png, or .gif. You can upload your picture to NYU Wordpress or imgur.com. However you want to do it, I just need a link here. This is required for HTML email to send."]);
  confPresets.push(["Mail Merge Feature Image URL (Enter \"No\" to send without pic, maximum height is 500 pixels):", "http://wp.nyu.edu/eric/wp-content/uploads/sites/3202/2017/07/laguardiacoop-small.png", "Customize the feature image for the email. Use Pretty HTML must be Yes for this to work. Has to be a link to a picture that ends in .jpg, .png, or .gif. You can upload your picture to NYU Wordpress or imgur.com. However you want to do it, I just need a link here, or say \"No\" for no image."]);
  confPresets.push(["Email Template:", "Hi {First Name} {Last Name},\n\nI want to contact you about an exciting new space that you may not have heard of. It is to my understanding that you are currently teaching {Class Name} in {School}.\n\nIf you have time, please stop by at the LaGuardia Co-op to learn about how you extend your classroom to the services we offer.\n\nBest regards,\n\nEric Lin", "Customize your email teplate. The replacement keywords are the column labels surrounded by braces.\n\nIf you have a column of First Names that you want to use, the top of that column must be the label \"First Names\", and the corresponding keyword would be {First Name}.\n\nYou don't have to use all the keywords, if the program can't find anything to replace, it will just simply ignore that keyword.\n\nIf you want to begin on a new line, you have to press Command + Enter (Ctrl + Enter on Windows) to begin a new line."]);
  confPresets.push(["Configuration Page Version (Do NOT Edit): ", "1", "Hit Mail Merge -> Process Mail Merge when everything is good to go"]);
  confData.setValues(confPresets);
  confData.setBackground("#ead1dc");
  conf.setColumnWidth(1, 400);
  conf.setColumnWidth(2, 500);
  conf.setColumnWidth(3, 600);
  confData.setWrap(true);
  confData = conf.getRange("B2:B15");
  confData.setBackground("#fff2cc");
  confData = conf.getRange("A1");
  confData.setFontSize(36);
  confData = conf.getRange("A15");
  confData.setFontSize(36);
  confData = conf.getRange("C1");
  confData.setFontSize(36);
  return "Configuration created successfully, I hope you're happy :)";
}

//Check if all the configuration is sort of correct. I tried to handle all the error cases, but there might be some I didn't think of.
function checkConfiguration() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("======== MAIL MERGE LOGS ========\n");
  Logger.log("----Starting Configuration Check----\n");
  
  Logger.log("Looking for mail merge configuration");
  //configuration sheet check
  var conf = ss.getSheetByName("Mail Merge Configuration");
  if(conf == null){
    return "Cannot find sheet named Mail Merge Configuration! This contains all my configuration items :(";
  }
  Logger.log("SUCCESS: Mail merge configuration found");
  
  //mail merge sheet check
  var name = conf.getRange(2, 2).getValue();
  Logger.log("Looking for sheet " + name + " as indicated");
  var sheet = ss.getSheetByName(name);
  if(sheet == null){
    return "Cannot find sheet " + name + " as indicated in the field named Mail Merge Active Sheet Name :(";
  }
  Logger.log("SUCCESS: Sheet named " + name + " was found");
  
  //mail merge email column check
  var colEmail = conf.getRange(3, 2).getValue();
  Logger.log("Looking for email colum as indicated");
  if(colEmail === "") {
    return "Cannot find the email column as indicated in the field named Mail Merge Email Column (Column Letter) :(";
  }
  Logger.log("SUCCESS: Email column " + colEmail + " was found");
  
  //mail merge data number check
  var numData = conf.getRange(4, 2).getValue()
  var numregex = /^\d+$/;
  Logger.log("Looking for data range as indicated");
  if(numData === "") {
    return "Cannot find the data column as indicated in the field named Mail Merge # of Data Columns :(";
  }
  else if(!numregex.test(numData)) {
    return "Not a valid number in the field named Mail Merge # of Data Columns :(";
  }
  Logger.log("SUCCESS: Data range " + numData + " was found");
  
  //mail merge send-as name check
  var sendName = conf.getRange(5, 2).getValue();
  Logger.log("Looking for Send As Name as indicated");
  if(sendName === ""){
    return "Cannot find Send As Name in the field named Mail Merge Send Mail From Name :(";
  }
  Logger.log("SUCCESS: Send name " + sendName + " was found");
  
  //mail merge send from email check
  Logger.log("Looking for send from email as indicated");
  var sendFromEmail = conf.getRange(6, 2).getValue();
  var emailregex = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  if(sendFromEmail === "") {
    return "Cannot find the send from email as indicated in the field named Mail Merge Send Mail From Email :(";
  }
  else if(!emailregex.test(sendFromEmail)) {
    return "You entered an invalid email in the field named Mail Merge Send Mail From Email :(";
  }
  var alias = GmailApp.getAliases();
  var aliasfound = false;
  for(var index = 0; index < alias.length; index++) {
    if(alias[index] === sendFromEmail) {
      aliasfound = true;
      break;
    }
  }
  if(!aliasfound) {
    return "Looks like you don't own the alias or it is not an acceptable alias to send from: " + sendFromEmail + " :(";
  }
  Logger.log("SUCCESS: Send from email " + sendFromEmail + " was found and validated");
  
  Logger.log("Looking for reply to email as indicated");
  var replyToEmail = conf.getRange(7, 2).getValue();
  if(replyToEmail === "") {
    return "Cannot find the reply to email as indicated in the field named Mail Merge Reply To :(";
  }
  else if(!emailregex.test(replyToEmail)) {
    return "You entered an invalid email in the field named Mail Merge Reply To :(";
  }
  Logger.log("SUCCESS: Reply to email " + replyToEmail + " was found");
  
  Logger.log("Looking for debug email as indicated");
  var debugEmail = conf.getRange(8, 2).getValue();
  if(debugEmail === "") {
    return "Cannot find the reply to email as indicated in the field named Mail Merge Reply To :(";
  }
  else if(!emailregex.test(debugEmail)) {
    return "You entered an invalid email in the field named Mail Merge Reply To :(";
  }
  Logger.log("SUCCESS: Debug email " + debugEmail + " was found");
  
  Logger.log("----Configuration Check Completed----\n");
  
  //mail merge get bcc conf
  var bccConf = conf.getRange(9, 2).getValue();
  var bcc = true;
  Logger.log("Looking for Email BCC Configuration as indicated");
  if(bccConf === ""){
    return "It seems like the Email BCC Configuration is empty :(";
  }
  else if(bccConf === "No") {
    bcc = false;
  }
  Logger.log("SUCCESS: Email BCC Configuration was found");
  
  var bcclist = conf.getRange(10, 2).getValue(); 
  if(bcc) {
    //mail merge get bcc list
    var bcclistcheck = bcclist.split(",");
    Logger.log("Looking for BCC Email List as indicated");
    for(var bindex = 0; bindex < bcclistcheck.length; bindex++) {
      if(!emailregex.test(bcclistcheck[bindex].trim()))
      {
        return "This is an invalid email to BCC to: " + bcclistcheck[bindex]; 
      }
    }
    Logger.log("SUCCESS: Email BCC List Configuration was found and tested");
  }
  
  //mail merge use pretty html
  var htmlConf = conf.getRange(11, 2).getValue();
  var useHtml = true;
  Logger.log("Looking for Pretty HTML Configuration as indicated");
  if(htmlConf === ""){
    return "It seems like the Use Pretty HTML Configuration is empty :(";
  }
  else if(htmlConf === "No") {
    useHtml = false;
  }
  Logger.log("SUCCESS: HTML Configuration was found");
  
  //mail merge get subject
  var subject = conf.getRange(12, 2).getValue();
  Logger.log("Looking for Email Subject Line as indicated");
  if(subject === ""){
    return "It seems like the Email Subject Line Configuration is empty :(";
  }
  Logger.log("SUCCESS: Email Subject Line was found");
  
  //mail merge feature image
  var headerimgurl = conf.getRange(13, 2).getValue();
  var imgurlregex = /\.(jpeg|jpg|gif|png)$/;
  Logger.log("Looking for header image url as indicated");
  if(headerimgurl === ""){
    return "It seems like the Header Image URL field is empty :(";
  }
  else if(!imgurlregex.test(headerimgurl)) {
    return "It seems like the Header Image URL is not a picture, this is a required field :(";
  }
  else {
    Logger.log("SUCCESS: Header Image URL was found");
  }
  
  //mail merge feature image
  var useimg = true;
  var imgurl = conf.getRange(14, 2).getValue();
  Logger.log("Looking for feature image url as indicated");
  if(imgurl === ""){
    return "It seems like the Feature Image URL field is empty :(";
  }
  else if(imgurl === "No") {
    useimg = false;
    Logger.log("SUCCESS: No feature image configured to be sent");
  }
  else if(!imgurlregex.test(imgurl)) {
    return "It seems like the Feature Image URL is not a picture :(";
  }
  else {
    Logger.log("SUCCESS: Feature Image URL was found");
  }
  
  
  //mail merge get template
  var template = conf.getRange(15, 2).getValue();
  Logger.log("Looking for Email Template as indicated");
  if(template === ""){
    return "It seems like the email template is empty :(";
  }
  Logger.log("SUCCESS: Email template was found");
  
  Logger.log("----Starting Mail Merge Email Validation----\n");
  var end = sheet.getLastRow();
  var colEmailNum = letterToColumn(colEmail);
  var emailList = sheet.getRange(2, colEmailNum, end - 1).getValues();
  for(var index = 0; index < emailList.length; index++) {
    Logger.log("Checking client email for validity: " + emailList[index][0]);
    if(!emailregex.test(emailList[index][0]))
    {
      Logger.log("Validity test FAILED: " + emailList[index][0]);
      mailLogs(debugEmail);
      return "There's an invalid email in " + name + " in column " + colEmail + ", row " + index + " (" + emailList[index][0] + "), please verify for accuracy!"; 
    }
    Logger.log("Validity test PASSED: " + emailList[index][0]);
  }
  Logger.log("----Mail Merge Email Validation Complete----\n");
  Logger.log("======== Starting Mail Merge Process ========\n");
  
  //Pass information to mailMerge method to start sending emails.
  mailMerge(sheet, colEmailNum, numData, sendName, sendFromEmail, replyToEmail, debugEmail, bcc, bcclist, useHtml, subject, headerimgurl, useimg, imgurl, template);
  Logger.log("Mail merge logs about to be sent to " + debugEmail);
  Logger.log("======== Mail Merge Process Completed ========\n");
  //Pass information to debugEmail for debugging purposes.
  mailLogs(sendFromEmail, sendName, replyToEmail, debugEmail);
  return "Thanks for confirming, your mail merge request has been submitted. A process log will be sent to " + debugEmail + " as indicated.";
}

//Converts column letters to a number
function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

//Converts column numbers into its corresponding letter
function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

//Grab mail merge data and loop over data row by row
function mailMerge(sheet, colEmailNum, numData, sendName, sendFromEmail, replyToEmail, debugEmail, bcc, bcclist, useHtml, subject, headerimgurl, useimg, imgurl, template) {
  var replacements = getLabelReplacements(sheet, numData);
  var endRow = sheet.getLastRow();
  var sheetData = sheet.getRange("A2:" + columnToLetter(numData) + endRow).getValues();
  for(var index = 0; index < sheetData.length; index++) {
    sendMailToClient(sheetData[index], colEmailNum, numData, sendName, sendFromEmail, replyToEmail, debugEmail, bcc, bcclist, useHtml, subject, headerimgurl, useimg, imgurl, template, replacements);
  }
}

//Replace template with client information, and send off email to client with/without bcc.
function sendMailToClient(data, colEmailNum, numData, sendName, sendFromEmail, replyToEmail, debugEmail, bcc, bcclist, useHtml, subject, headerimgurl, useimg, imgurl, template, replacements) {
  String.prototype.replaceAll = function(target, replacement) {
    return this.split(target).join(replacement);
  };
  Logger.log("Generating email to " + data[colEmailNum - 1]);
  var message = template;
  for(var index = 0; index < replacements.length; index++) {
    message = message.replaceAll(replacements[index], data[index]);
  }
  Logger.log("Inserted merge details to template " + data[colEmailNum - 1]);
  
  if(useHtml) {
    message = message.replace(/(?:\r\n|\r|\n)/g, '<br />');
    Logger.log("Converted template to HTML body");
    
    var img = "";
    if(useimg) {
      img = "<img src=\"" + imgurl + "\" alt=\"Insert image here\" title=\"Insert image here\" style=\"outline: none;text-decoration: none;-ms-interpolation-mode: bicubic;clear: both;display: block;border: 0;height: auto;float: none;width: 100%; margin: auto; max-width: 500px;\" width=\"500\"/>";
      Logger.log("Inserted image into HTML " + data[colEmailNum - 1]);
    }
  
    //HTML TEMPLATE, DO NOT CHANGE UNLESS YOU KNOW WHAT YOU'RE DOING.
    var message = "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional //EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">" + 
    "<!--[if IE]><html xmlns=\"http://www.w3.org/1999/xhtml\" class=\"ie-browser\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:" + 
    "office:office\"><![endif]--><!--[if !IE]><!--><html style=\"margin: 0;padding: 0;\" xmlns=\"http://www.w3.org/1999/xhtml\"><!--<![endif]--><head><!--[if gte mso" +
    "9]><xml><o:OfficeDocumentSettings><o:AllowPNG/><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml><![endif]--><meta http-equiv=\"Content-Type\"" +
    "content=\"text/html; charset=utf-8\"><meta name=\"viewport\" content=\"width=device-width\"><!--[if !mso]><!--><meta http-equiv=\"X-UA-Compatible\" content=\"" +
    "IE=edge\"><!--<![endif]--><style type=\"text/css\" id=\"media-query\">body {margin: 0;padding: 0; }table {border-collapse: collapse;table-layout: fixed; }" +
    "* {line-height: inherit; }a[x-apple-data-detectors=true] {color: inherit !important;text-decoration: none !important; }[owa] .img-container div, [owa] .img-container" +
    "button {display: block !important; }[owa] .fullwidth button {width: 100% !important; }.ie-browser .col, [owa] .block-grid .col {display: table-cell;float: none" +
    "!important;vertical-align: top; }.ie-browser .num12, .ie-browser .block-grid, [owa] .num12, [owa] .block-grid {width: 500px !important; }.ExternalClass," +
    ".ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div {line-height: 100%; }.ie-browser .mixed-two-up .num4, [owa]" +
    ".mixed-two-up .num4 {width: 164px !important; }.ie-browser .mixed-two-up .num8, [owa] .mixed-two-up .num8 {width: 328px !important; }.ie-browser .block-grid.two-up" +
    ".col, [owa] .block-grid.two-up .col {width: 250px !important; }.ie-browser .block-grid.three-up .col, [owa] .block-grid.three-up .col {width: 166px !important; }" +
    ".ie-browser .block-grid.four-up .col, [owa] .block-grid.four-up .col {width: 125px !important; }.ie-browser .block-grid.five-up .col, [owa] .block-grid.five-up" +
    ".col {width: 100px !important; }.ie-browser .block-grid.six-up .col, [owa] .block-grid.six-up .col {width: 83px !important; }.ie-browser .block-grid.seven-up" +
    ".col, [owa] .block-grid.seven-up .col {width: 71px !important; }.ie-browser .block-grid.eight-up .col, [owa] .block-grid.eight-up .col {width: 62px !important; }" +
    ".ie-browser .block-grid.nine-up .col, [owa] .block-grid.nine-up .col {width: 55px !important; }.ie-browser .block-grid.ten-up .col, [owa] .block-grid.ten-up .col {" +
    "width: 50px !important; }.ie-browser .block-grid.eleven-up .col, [owa] .block-grid.eleven-up .col {width: 45px !important; }.ie-browser .block-grid.twelve-up .col," +
    "[owa] .block-grid.twelve-up .col {width: 41px !important; }@media only screen and (min-width: 520px) {.block-grid {width: 500px !important; }.block-grid .col {" +
    "display: table-cell;Float: none !important;vertical-align: top; }.block-grid .col.num12 {width: 500px !important; }.block-grid.mixed-two-up .col.num4 {width: 164px" +
    "!important; }.block-grid.mixed-two-up .col.num8 {width: 328px !important; }.block-grid.two-up .col {width: 250px !important; }.block-grid.three-up .col {" +
    "width: 166px !important; }.block-grid.four-up .col {width: 125px !important; }.block-grid.five-up .col {width: 100px !important; }.block-grid.six-up .col {width:" +
    "83px !important; }.block-grid.seven-up .col {width: 71px !important; }.block-grid.eight-up .col {width: 62px !important; }.block-grid.nine-up .col {width: 55px" +
    "!important; }.block-grid.ten-up .col {width: 50px !important; }.block-grid.eleven-up .col {width: 45px !important; }.block-grid.twelve-up .col {width: 41px !important; } }" +
    "@media (max-width : 320px ){.block-grid, .col {min-width: 320px !important;max-width: 100% !important; }.block-grid {width: calc(100% - 40px) !important; }.col {" +
    "width: 100% !important; }.col > div {margin: 0 auto; }img.fullwidth {max-width: 100% !important; } }</style></head><!--[if mso]><body class=\"mso-container\"" +
    "style=\"background-color:#FFFFFF;\"><![endif]--><!--[if !mso]><!--><body class=\"clean-body\" style=\"margin: 0;padding: 0;-webkit-text-size-adjust: 100%;" +
    "background-color: #FFFFFF\"><!--<![endif]--><div class=\"nl-container\" style=\"min-width: 320px;Margin: 0 auto;background-color: #FFFFFF\"><!--[if (mso)|(IE)]>" +
    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td align=\"center\" style=\"background-color: #FFFFFF;\"><![endif]--><div style=\"" +
    "background-color:#F9F9F9;\"><div style=\"Margin: 0 auto;min-width: 320px;max-width: 500px;width: 500px;width: calc(19000% - 98300px);overflow-wrap: break-word;" +
    "word-wrap: break-word;word-break: break-word;background-color: transparent;\" class=\"block-grid \"><div style=\"border-collapse: collapse;display: table;width:" +
    "100%;\"><!--[if (mso)|(IE)]><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td style=\"background-color:#F9F9F9;\" align=\"center\">" +
    "<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"width: 500px;\"><tr class=\"layout-full-width\" style=\"background-color:transparent;\"><![endif]-->" +
    "<!--[if (mso)|(IE)]><td align=\"center\" width=\"500\" style=\" width:500px; padding-right: 0px; padding-left: 0px; border-top: 0px solid transparent; border-left:" +
    "0px solid transparent; border-bottom: 0px solid transparent; border-right: 0px solid transparent;\" valign=\"top\"><![endif]--><div class=\"col num12\" style=\"min-width:" +
    "320px;max-width: 500px;width: 500px;width: calc(18000% - 89500px);background-color: transparent;\"><!--[if (!mso)&(!IE)]><!--><div style=\"border-top: 0px solid" +
    "transparent; border-left: 0px solid transparent; border-bottom: 0px solid transparent; border-right: 0px solid transparent;\"><!--<![endif]--><div style=\"background-color:" +
    "transparent; display: inline-block!important; width: 100% !important;\"><div style=\"Margin-top:0px; Margin-bottom:0px;\"><div align=\"center\" class=\"img-container" +
    "center fullwidth\"><!--[if !mso]><!--><div style=\"Margin-right: 0px;Margin-left: 0px;\"><!--<![endif]--><!--[if mso]><table width=\"500\" cellpadding=\"0\"" +
    "cellspacing=\"0\" border=\"0\"><tr><td style=\"padding-right:  0px; padding-left: 0px;\" align=\"center\"><![endif]--><div align=\"center\" style=\"margin-left:" +
    "15px; margin-right: 15px; padding-top: 30px\"><img class=\"center fullwidth\" align=\"center\" border=\"0\" src=\"" + headerimgurl + "\"" +
    "alt=\"Image\" title=\"Image\" style=\"outline: none;text-decoration: none;-ms-interpolation-mode: bicubic;clear: both;display: block;border: 0;height: auto;float:" +
    "none;width: 100%; margin: auto; max-width: 500px;\" width=\"500\"></div><!--[if mso]></td></tr></table><![endif]--><div style=\"line-height:15px;font-size:1px\">&nbsp;</div>" +
    "<!--[if !mso]><!--></div><!--<![endif]--></div><!--[if !mso]><!--><div style=\"Margin-right: 15px; Margin-left: 15px;\"><!--<![endif]--><div style=\"line-height: 10px;" +
    "font-size: 1px\">&nbsp;</div><!--[if mso]><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td style=\"padding-right: 0px; padding-left: 0px;\"><![endif]-->" +
    "<div align=\"center\" style=\"margin-bottom: 15px\">" + img +
    "</div><div style=\"font-size:12px;line-height:14px;color:#555555;font-family:Arial, 'Helvetica Neue', Helvetica, sans-serif;text-align:left;\">" + message + "</div><!--[if mso]></td></tr></table><![endif]--><div style=\"line-height: 10px;" +
    "font-size: 1px\">&nbsp;</div><!--[if !mso]><!--></div><!--<![endif]--><!--[if !mso]><!--><div align=\"center\" style=\"Margin-right: 10px;Margin-left: 10px;\">" +
    "<!--<![endif]--><div style=\"line-height: 10px; font-size:1px\">&nbsp;</div><!--[if (mso)|(IE)]><table width=\"100%\" align=\"center\" cellpadding=\"0\" cellspacing=\"" +
    "0\" border=\"0\"><tr><td style=\"padding-right: 10px;padding-left: 10px;\"><![endif]--><div style=\"border-top: 10px solid transparent; width:100%; font-size:1px;\"" +
    ">&nbsp;</div><!--[if (mso)|(IE)]></td></tr></table><![endif]--><div style=\"line-height:10px; font-size:1px\">&nbsp;</div><!--[if !mso]><!--></div><!--<![endif]-->" +
    "<!--[if !mso]><!--><div align=\"center\" style=\"Margin-right: 0px;Margin-left: 0px;\"><!--<![endif]--><!--[if (mso)|(IE)]><table width=\"100%\" align=\"center\"" +
    "cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td style=\"padding-right: 0px;padding-left: 0px;\"><![endif]--><div style=\"border-top: 1px solid #BBBBBB;" +
    "width:100%; font-size:1px;\">&nbsp;</div><!--[if (mso)|(IE)]></td></tr></table><![endif]--><!--[if !mso]><!--></div><!--<![endif]--></div></div><!--[if (!mso)&(!IE)]>" +
    "<!--></div><!--<![endif]--></div><!--[if (mso)|(IE)]></tr></table></td></tr></table><![endif]--></div></div></div><div style=\"background-color:#F9F9F9;\">" +
    "<div style=\"Margin: 0 auto;min-width: 320px;max-width: 500px;width: 500px;width: calc(19000% - 98300px);overflow-wrap: break-word;word-wrap: break-word;word-break:" +
    "break-word;background-color: transparent;\" class=\"block-grid \"><div style=\"border-collapse: collapse;display: table;width: 100%;\"><!--[if (mso)|(IE)]><table" +
    "width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td style=\"background-color:#F9F9F9;\" align=\"center\"><table cellpadding=\"0\" cellspacing=\"" +
    "0\" border=\"0\" style=\"width: 500px;\"><tr class=\"layout-full-width\" style=\"background-color:transparent;\"><![endif]--><!--[if (mso)|(IE)]><td align=\"center\"" +
    "width=\"500\" style=\" width:500px; padding-right: 0px; padding-left: 0px; border-top: 0px solid transparent; border-left: 0px solid transparent; border-bottom: 0px solid" +
    "transparent; border-right: 0px solid transparent;\" valign=\"top\"><![endif]--><div class=\"col num12\" style=\"min-width: 320px;max-width: 500px;width: 500px;" +
    "width: calc(18000% - 89500px);background-color: transparent;\"><!--[if (!mso)&(!IE)]><!--><div style=\"border-top: 0px solid transparent; border-left: 0px solid" +
    "transparent; border-bottom: 0px solid transparent; border-right: 0px solid transparent;\"><!--<![endif]--><div style=\"background-color: transparent; display:" +
    "inline-block!important; width: 100% !important;\"><div style=\"Margin-top:15px; Margin-bottom:30px;\"><div align=\"center\" class=\"img-container center fullwidth\"" +
    "style=\"margin-left: 15px; margin-right: 15px\"><!--[if !mso]><!--><div style=\"Margin-right: 0px;Margin-left: 0px; margin-top: 0px;\"><!--<![endif]--><!--[if mso]>" +
    "<table width=\"500\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td style=\"padding-right:  0px; padding-left: 0px;\" align=\"center\"><![endif]-->" +
    "<a href=\"http://www.nyu.edu/life/information-technology/locations-and-facilities/student-technology-centers/laguardia-co-op.html\" target=\"_blank\">" +
    "<img class=\"center fullwidth\" align=\"center\" border=\"0\" src=\"http://wp.nyu.edu/jl4884/wp-content/uploads/sites/6605/2017/04/NYUIT.png\" alt=\"Image\" title=\"" +
    "Image\" style=\"outline: none;text-decoration: none;-ms-interpolation-mode: bicubic;clear: both;display: block;border: none;padding-top:none;margin-top:none;height: auto;" +
    "float: none;width: 100%;max-width: 500px\" width=\"500\"></a><!--[if mso]></td></tr></table><![endif]--><!--[if !mso]><!--></div><!--<![endif]--></div>" +
    "<!--[if !mso]><!--><div style=\"Margin-right: 10px; Margin-left: 0px;\"><!--<![endif]--><!--[if mso]><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\"" +
    "border=\"0\"><tr><td style=\"padding-right: 10px; padding-left: 0px;\"><![endif]--><div style=\"font-size:12px;line-height:14px;color:#555555;font-family:Arial," +
    "'Helvetica Neue', Helvetica, sans-serif;text-align:left; margin-left: 15px; margin-right: 15px\"><p style=\"margin: 0; padding-top: 5px;font-size: 12px;line-height:" +
    "16px\"><span style=\"color: rgb(0, 141, 150); font-size: 12px; line-height: 14px;\"><em><strong>Connect &amp; Collaborate at the LaGuardia Co-op</strong></em></span>" +
    "</p><p style=\"margin: 0;font-size: 12px;line-height: 16px\"><span style=\"color: rgb(137, 137, 137); font-size: 12px; line-height: 14px;\"><em>539-541 LaGuardia Place," +
    "New York NY 10012 </em></span><br><span style=\"color: rgb(137, 137, 137); font-size: 12px; line-height: 14px;\"><em>(212) 998-3427</em></span></p></div><!--[if mso]>" +
    "</td></tr></table><![endif]--><div style=\"line-height: 10px; font-size: 1px\">&nbsp;</div><!--[if !mso]><!--></div><!--<![endif]--><div align=\"center\"" +
    "style=\"Margin-right: 10px; Margin-left: 10px; Margin-bottom: 10px;\"><div style=\"line-height:10px;font-size:1px\">&nbsp;</div><div style=\"display: table;" +
    "max-width:131px;\"><!--[if (mso)|(IE)]><table width=\"131\" align=\"center\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"border-collapse:collapse;" +
    "mso-table-lspace: 0pt;mso-table-rspace: 0pt; width:131px;\"><tr><td width=\"37\" style=\"width:37px;\" valign=\"top\"><![endif]--><table align=\"left\" border=\"0\"" +
    "cellspacing=\"0\" cellpadding=\"0\" width=\"32\" height=\"32\" style=\"border-spacing: 0;border-collapse: collapse;mso-table-lspace: 0pt;mso-table-rspace: 0pt;" +
    "vertical-align: top;Margin-right: 5px\"><tbody><tr style=\"vertical-align: top\"><td align=\"left\" valign=\"middle\" style=\"word-break: break-word;border-collapse:" +
    "collapse !important;vertical-align: top\"><a href=\"https://www.facebook.com/nyustc/\" title=\"Facebook\" target=\"_blank\"><img src=\"http://wp.nyu.edu/jl4884/wp-content" +
    "/uploads/sites/6605/2017/04/facebook@2x.png\" alt=\"Facebook\" title=\"Facebook\" width=\"32\" style=\"outline: none;text-decoration: none;-ms-interpolation-mode:" +
    "bicubic;clear: both;display: block;border: none;height: auto;float: none;max-width: 32px !important\"></a><div style=\"line-height:5px;font-size:1px\">&nbsp;</div>" +
    "</td></tr></tbody></table><!--[if (mso)|(IE)]></td><td width=\"37\" style=\"width:37px;\" valign=\"top\"><![endif]--><table align=\"left\" border=\"0\" cellspacing=\"0\"" +
    "cellpadding=\"0\" width=\"32\" height=\"32\" style=\"border-spacing: 0;border-collapse: collapse;mso-table-lspace: 0pt;mso-table-rspace: 0pt;vertical-align: top;" +
    "Margin-right: 5px\"><tbody><tr style=\"vertical-align: top\"><td align=\"left\" valign=\"middle\" style=\"word-break: break-word;border-collapse: collapse !important;" +
    "vertical-align: top\"><a href=\"https://www.instagram.com/nyustc/\" title=\"Instagram\" target=\"_blank\"><img src=\"http://wp.nyu.edu/jl4884/wp-content/uploads/" +
    "sites/6605/2017/04/instagram@2x.png\" alt=\"Instagram\" title=\"Instagram\" width=\"32\" style=\"outline: none;text-decoration: none;-ms-interpolation-mode: bicubic;" +
    "clear: both;display: block;border: none;height: auto;float: none;max-width: 32px !important\"></a><div style=\"line-height:5px;font-size:1px\">&nbsp;</div></td>" +
    "</tr></tbody></table><!--[if (mso)|(IE)]></td><td width=\"37\" style=\"width:37px;\" valign=\"top\"><![endif]--><table align=\"left\" border=\"0\" cellspacing=\"0\"" +
    "cellpadding=\"0\" width=\"32\" height=\"32\" style=\"border-spacing: 0;border-collapse: collapse;mso-table-lspace: 0pt;mso-table-rspace: 0pt;vertical-align: top;" +
    "Margin-right: 0\"><tbody><tr style=\"vertical-align: top\"><td align=\"left\" valign=\"middle\" style=\"word-break: break-word;border-collapse: collapse !important;" +
    "vertical-align: top\"><a href=\"http://www.nyu.edu/life/information-technology/locations-and-facilities/student-technology-centers/laguardia-co-op.html\" title=\"Web" +
    "Site\" target=\"_blank\"><img src=\"http://wp.nyu.edu/jl4884/wp-content/uploads/sites/6605/2017/04/website@2x.png\" alt=\"Web Site\" title=\"Web Site\" width=\"32\"" +
    "style=\"outline: none;text-decoration: none;-ms-interpolation-mode: bicubic;clear: both;display: block;border: none;height: auto;float: none;max-width: 32px !important\">" +
    "</a><div style=\"line-height:5px;font-size:1px\">&nbsp;</div></td></tr></tbody></table><!--[if (mso)|(IE)]></td></tr></table><table width=\"100%\" cellpadding=\"0\"" +
    "cellspacing=\"0\" border=\"0\"><tr><td>&nbsp;</td></tr></table><![endif]--></div></div><!--[if !mso]><!--><div align=\"center\" style=\"Margin-right: 10px;" +
    "Margin-left: 10px;\"><!--<![endif]--><div style=\"line-height: 10px; font-size:1px\">&nbsp;</div><!--[if (mso)|(IE)]><table width=\"100%\" align=\"center\"" +
    "cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td style=\"padding-right: 10px;padding-left: 10px;\"><![endif]--><div style=\"border-top: 0px solid transparent;" +
    "width:100%; font-size:1px;\">&nbsp;</div><!--[if (mso)|(IE)]></td></tr></table><![endif]--><div style=\"line-height:10px; font-size:1px\">&nbsp;</div><!--[if !mso]><!-->" +
    "</div><!--<![endif]--></div></div><!--[if (!mso)&(!IE)]><!--></div><!--<![endif]--></div><!--[if (mso)|(IE)]></tr></table></td></tr></table><![endif]--></div>" +
    "</div></div><!--[if (mso)|(IE)]></td></tr></table><![endif]--></div></body></html>";
  
    Logger.log("Inserted details into HTML for " + data[colEmailNum - 1]);
  }
  
  var alias = GmailApp.getAliases();
  var aliasIndex = 0;
  for(var index = 0; index < alias.length; index++) {
    if(alias[index] === sendFromEmail) {
      aliasIndex = index;
      break;
    }
  }
  
  if(bcc && useHtml) {
    GmailApp.sendEmail(
      data[colEmailNum - 1],         
      subject,               
      message, {           
        from: GmailApp.getAliases()[aliasIndex],
        name: sendName,
        replyTo: replyToEmail,
        bcc: bcclist,
        htmlBody: message
      }
    ); 
    Logger.log("Email has been sent with HTML to " + data[colEmailNum - 1] + " with BCC to " + bcclist + "\n");
  }
  else if(!bcc && useHtml){
    GmailApp.sendEmail(
      data[colEmailNum],         
      subject,               
      message, {           
        from: GmailApp.getAliases()[aliasIndex],
        name: sendName,
        replyTo: replyToEmail,
        htmlBody: message
      }
    ); 
    Logger.log("Email has been sent with HTML to " + data[colEmailNum] + "\n");
  }
  else if(bcc && !useHtml) {
    GmailApp.sendEmail(
      data[colEmailNum - 1],         
      subject,               
      message, {           
        from: GmailApp.getAliases()[aliasIndex],
        name: sendName,
        replyTo: replyToEmail,
        bcc: bcclist
      }
    ); 
    Logger.log("Email has been sent without HTML to " + data[colEmailNum - 1] + " with BCC to " + bcclist + "\n");
  }
  else {
    GmailApp.sendEmail(
      data[colEmailNum],         
      subject,               
      message, {           
        from: GmailApp.getAliases()[aliasIndex],
        name: sendName,
        replyTo: replyToEmail
      }
    ); 
    Logger.log("Email has been sent without HTML to " + data[colEmailNum] + "\n");
  }
  
}

//Processes the keyword replacements
function getLabelReplacements(sheet, numData) {
  Logger.log("----Generating Keywords----\n")
  var labels = sheet.getRange(1, 1, 1, numData).getValues()[0];
  var labelReplacements = [];
  for(var index = 0; index < labels.length; index++) {
    Logger.log("{" + labels[index] + "}");
    labelReplacements.push("{" + labels[index] + "}");
  }
  Logger.log("----Keywords Generated----\n")
  return labelReplacements;
}

//Email debug logs to debugEmail
function mailLogs(sendFromEmail, sendName, replyToEmail, debugEmail) {
  var subject = "Mail Merge Logs - " + Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm:ss");
  var alias = GmailApp.getAliases();
  var aliasIndex = 0;
  for(var index = 0; index < alias.length; index++) {
    if(alias[index] === sendFromEmail) {
      aliasIndex = index;
      break;
    }
  }
  var message = Logger.getLog();
  GmailApp.sendEmail(
    debugEmail,         
    subject,               
    message, {           
      from: GmailApp.getAliases()[aliasIndex],
      name: sendName,
      replyTo: replyToEmail,
    }
  ); 
}

