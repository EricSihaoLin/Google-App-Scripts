/* Giphy Email
 * By Eric Lin
 *
 * This script SHOULD live in a Google Spreadsheet's Script Editor
 * To access the script editor: Tools -> Script Editor
 * and then copy and paste this entire document into the code editor. Press save, then close out the script editor.
 * Close and reopen the document, the document should have all the features installed upon relaunch
 * 
 * This script requires into to the 3 essential columns, the Giphy url that ends in .gif extension, the NetID of the client, and the name of the client.
 * 
 * Set the cronJob function to run every minute. Make sure this script is running on the account you desire that has your departmental email.
 */

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Input Data')
      .addItem('New Giphy Data', 'newData')
      .addToUi();
}

function newData() {
  var html = HtmlService.createHtmlOutputFromFile('index').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi()
       .showModalDialog(html, 'Add Entry');
}

function itemAdd(form) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("New GIF");
  sheet.appendRow([form.url, form.netid, form.name]);
  return true;
}

function cronJobGiphy() {
   
  //spreadsheet variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var newGif = ss.getSheetByName("New GIF");
  var processed = ss.getSheetByName("Processed GIF");
  var numnewGifProcessed = 0;
  var pendingGifs = newGif.getLastRow();
  
  //our title row counts for 1, have to disregard that
  if(pendingGifs != 1)
  {
    var currentGif = 2;
    //FIRST RUN: Get information and email client then send job run notification to admin
    Logger.log("Number of GIFs to process: " + (pendingGifs - 1));
    while(currentGif <= pendingGifs)
    {
      var range = newGif.getRange("A" + currentGif + ":C" + currentGif);
      var data = range.getValues();
      
      Logger.log("Sending GIF " + (currentGif - 1) + " to respective client");
      pushEmails(data);
      Logger.log("GIF " + (currentGif - 1) + " has been sent");
      
      numnewGifProcessed++;
      currentGif++;
    }
    
    //SECOND RUN: Move the booking information to archive sheet
    var range = newGif.getRange("A2:C" + pendingGifs);
    var data = range.getValues();
    for(var i = 0; i < data.length; i++)
    {
      processed.appendRow(data[i]); 
    }
    newGif.deleteRows(2, data.length);
  }
  
  Logger.log("GIFs finished processing: " + numnewGifProcessed);
  
  if(numnewGifProcessed > 0)
  {
    var emailAddress = "eric.lin@nyu.edu";
    var subject = "Giphy Logs - " + new Date(); 
    var message = Logger.getLog();
    MailApp.sendEmail(emailAddress, subject, message);
  }
}

function pushEmails(data)
{
  var url = data[0][0];
  var email = data[0][1] + "@nyu.edu";
  var name = data[0][2];
  
  if(url || email || name) {
    Logger.log("One of the essential field was empty, skipping! URL: " + url + ", Email: " + email + ", Name: " + name);
    return;
  }
  
  var title = "Your ✨NEW✨ GIF is Ready " + name + "!";
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
    "15px; margin-right: 15px; padding-top: 30px\"><img class=\"center fullwidth\" align=\"center\" border=\"0\" src=\"http://wp.nyu.edu/eric/wp-content/uploads/sites/3202/2017/07/Header-1.png\"" +
    "alt=\"Image\" title=\"Image\" style=\"outline: none;text-decoration: none;-ms-interpolation-mode: bicubic;clear: both;display: block;border: 0;height: auto;float:" +
    "none;width: 100%; margin: auto; max-width: 500px;\" width=\"500\"></div><!--[if mso]></td></tr></table><![endif]--><div style=\"line-height:15px;font-size:1px\">&nbsp;</div>" +
    "<!--[if !mso]><!--></div><!--<![endif]--></div><!--[if !mso]><!--><div style=\"Margin-right: 15px; Margin-left: 15px;\"><!--<![endif]--><div style=\"line-height: 10px;" +
    "font-size: 1px\">&nbsp;</div><!--[if mso]><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td style=\"padding-right: 0px; padding-left: 0px;\"><![endif]-->" +
    "<div align=\"center\" style=\"margin-bottom: 15px\"><img src=\"" + url + "\" alt=\"Insert image here\" title=\"Insert image here\" style=\"outline: none;text-decoration:" +
    "none;-ms-interpolation-mode: bicubic;clear: both;display: block;border: 0;height: auto;float: none;width: 100%; margin: auto; max-width: 500px;\" width=\"500\"/>" +
    "</div><div style=\"font-size:12px;line-height:14px;color:#555555;font-family:Arial, 'Helvetica Neue', Helvetica, sans-serif;text-align:left;\"><p style=\"margin: 0;" +
    "font-size: 12px;line-height: 14px;text-align: justify\">Thanks for visiting our Video Recording Booth " + name + "! Look at you, now you're immortalized in GIF form. You shall have" +
    " the power to brag about this to all your friends!</p><br><p style=\"margin: 0;font-size: 12px;line-height: 14px;text-align: justify\"><strong>Don't forget to come back" +
    " to LaGuardia Co-op during the academic year and keep up with our upcoming workshops and events by <a href=\"https://groups.google.com/a/nyu.edu/forum/?hl=en#!forum/laguardia.co-op/join\">" +
    "subscribing to our mailing list!</a></strong></p></div><!--[if mso]></td></tr></table><![endif]--><div style=\"line-height: 10px;" +
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
  
  var sender = GmailApp.getAliases()[0] + "";
  
  GmailApp.sendEmail(
    email,         
    title,               
    'test', {           
      from: sender,
      name: 'LaGuardia Co-op',
      replyTo: 'student.tech.centers@nyu.edu',
      htmlBody: message
    }
  ); 
  Logger.log("Email notification successfully created for client " + name);
  
}