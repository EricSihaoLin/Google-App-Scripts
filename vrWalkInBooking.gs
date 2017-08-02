/* VR Walk In Booking
 * By Eric Lin
 *
 * This script SHOULD live in a Google Spreadsheet's Script Editor
 * To access the script editor: Tools -> Script Editor
 * and then copy and paste this entire document into the code editor. Press save, then close out the script editor.
 * Close and reopen the document, the document should have all the features installed upon relaunch
 * 
 * This script requires a form attached to a spreadsheet, with the following result columns: Timestamp, Email Address, Your Reservation Date, Your Reservation Duration (in Minutes), Your First & Last Name, Your Email, Your NetID, Your Phone Number, Expected Number of Attendees, Console Option, Usage Reason, Please explain briefly about the nature of your project, we'll love to hear what you're working on!, Will any Non-NYU individuals be present during the session?, If yes, list the expected non-NYU guest here, Please indicate that you have read our Terms of Use
 * 
 * Set the cronJob function to run every minute. Make sure this script is running on the account you desire that has your departmental email.
 */

function cronJobBookingForm() {
   
  //spreadsheet variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bookings = ss.getSheetByName("Bookings");
  var processed = ss.getSheetByName("Processed Bookings");
  var numBookingsProcessed = 0;
  var pendingBookings = bookings.getLastRow();
  
  //calendar variables
  var calendar = CalendarApp.getCalendarById('nyu.edu_pshid3c02pqduechcf9a3hjbt8@group.calendar.google.com');
  
  //our title row counts for 1, have to disregard that
  if(pendingBookings != 1)
  {
    var currentBooking = 2;
    //FIRST RUN: Get information and create calendar event then send email notification
    Logger.log("Number of bookings: " + (pendingBookings - 1));
    while(currentBooking <= pendingBookings)
    {
      var range = bookings.getRange("A" + currentBooking + ":O" + currentBooking);
      var data = range.getValues();
      
      Logger.log("Pushing Booking " + (currentBooking - 1) + " to Calendar");
      var tempID = pushToCalendar(data, calendar);
      bookings.getRange("P" + currentBooking + ":P" + currentBooking).setValue(tempID);
      Logger.log("Calendar event created, ID: " + tempID);
      
      Logger.log("Sending Emails for Booking " + (currentBooking - 1));
      pushEmails(data);
      
      numBookingsProcessed++;
      currentBooking++;
    }
    
    //SECOND RUN: Move the booking information to archive sheet
    var range = bookings.getRange("A2:P" + pendingBookings);
    var data = range.getValues();
    for(var i = 0; i < data.length; i++)
    {
      processed.appendRow(data[i]); 
    }
    bookings.deleteRows(2, data.length);
  }
  
  Logger.log("Bookings processed: " + numBookingsProcessed);
  
  if(numBookingsProcessed > 0)
  {
    var emailAddress = "eric.lin@nyu.edu";
    var subject = "VR Walk-in Booking Process Logs - " + new Date(); 
    var message = Logger.getLog();
    MailApp.sendEmail(emailAddress, subject, message);
  }
}

function cronJobPullEvents()
{
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(0,0,0,0);
  var today = new Date();
  today.setHours(0,0,0,0);
  
  var calendar = CalendarApp.getCalendarById('nyu.edu_pshid3c02pqduechcf9a3hjbt8@group.calendar.google.com');
  var events = calendar.getEvents(yesterday, today,{search: '-Walk-in'});
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var processed = ss.getSheetByName("Processed Bookings");
  var start = 2;
  var end = processed.getLastRow();
  if(!(end < 501)){
    start = end - 500;
  }
  var last500 = processed.getRange(start, 16, end - 1).getValues();
  
  Date.prototype.formatDate = function (format) {
    var date = this,
        day = date.getDate(),
        month = date.getMonth() + 1,
        year = date.getFullYear(),
        hours = date.getHours(),
        minutes = date.getMinutes(),
        seconds = date.getSeconds();

    if (!format) {
        format = "MM/dd/yyyy";
    }

    format = format.replace("MM", month.toString().replace(/^(\d)$/, '0$1'));

    if (format.indexOf("yyyy") > -1) {
        format = format.replace("yyyy", year.toString());
    } else if (format.indexOf("yy") > -1) {
        format = format.replace("yy", year.toString().substr(2, 2));
    }

    format = format.replace("dd", day.toString().replace(/^(\d)$/, '0$1'));

    if (format.indexOf("t") > -1) {
        if (hours > 11) {
            format = format.replace("t", "pm");
        } else {
            format = format.replace("t", "am");
        }
    }

    if (format.indexOf("HH") > -1) {
        format = format.replace("HH", hours.toString().replace(/^(\d)$/, '0$1'));
    }

    if (format.indexOf("hh") > -1) {
        if (hours > 12) {
            hours -= 12;
        }

        if (hours === 0) {
            hours = 12;
        }
        format = format.replace("hh", hours.toString().replace(/^(\d)$/, '0$1'));
    }

    if (format.indexOf("mm") > -1) {
        format = format.replace("mm", minutes.toString().replace(/^(\d)$/, '0$1'));
    }

    if (format.indexOf("ss") > -1) {
        format = format.replace("ss", seconds.toString().replace(/^(\d)$/, '0$1'));
    }

    return format;
  };
  
  
  for(var i = 0; i < events.length; i++)
  {
    var eventID = events[i].getId();
    var found = false;
    for(var j = 0; j < last500.length; j++) {
      if(eventID === last500[j][0]) {
        found = true;
      }
    }
    if(!found){
      var timecreated = events[i].getDateCreated().formatDate('MM/dd/yyyy HH:mm:ss');
      var details = events[i].getDescription();
      var timereservation = events[i].getStartTime().formatDate('MM/dd/yyyy HH:mm:ss');
      var duration = Math.floor((Math.abs(events[i].getEndTime() - events[i].getStartTime())/1000)/60);
      var clientName = extractClientName(details);
      var clientEmail = extractClientEmail(details);
      var clientNetID = extractClientNetID(details);
      var clientPhone = extractClientPhone(details);
      var clientNumGuest = extractClientNumGuest(details);
      var clientConsoleOption = extractClientConsoleOption(details);
      var clientUsage = extractClientUsage(details);
      var clientUsageExplanation = extractClientUsageExplanation(details);
      var clientNonNYUGuest = extractClientNonNYUGuest(details);
      var clientNonNYUGuestList = extractClientNonNYUGuestList(details);
      var clientAgreedToU = extractClientAgreedToU(details);
      processed.appendRow([timecreated, clientEmail, timereservation, duration, clientName, clientEmail, clientNetID, clientPhone, clientNumGuest, clientConsoleOption, clientUsage, clientUsageExplanation, clientNonNYUGuest, clientNonNYUGuestList, clientAgreedToU, eventID]);
    }
  }
}

function pushToCalendar(data, calendar) {
  var startTime = new Date(data[0][2]);
  var duration = data[0][3];
  var endTime = new Date(startTime.getTime() + data[0][3] * 60 * 1000);
  var clientName = data[0][4];
  var clientEmail = data[0][5];
  var clientNetID = data[0][6];
  var clientPhone = data[0][7];
  var clientNumGuest = data[0][8];
  var clientConsoleOption = data[0][9];
  var clientUsage = data[0][10];
  var clientUsageExplanation = data[0][11];
  var clientNonNYUGuest = data[0][12];
  var clientNonNYUGuestList = data[0][13];
  var clientAgreedToU = data[0][14];
      
  var title = clientName + " (Walk-in) - " + clientConsoleOption + " - " + duration + " Minutes";
  var location = clientPhone + "";
  var description = "Client:\n" +
        clientName + " (" + clientEmail + ")\n\n" + "Client NetID:\n" + clientNetID + "\n\n" +
          "Duration:\n" + clientConsoleOption + " - " + duration + " Minutes\n\n" +
            "Booking page:\n" + "Virtual Reality Booking Page (Walk-in)\n\n" +
              "Phone:\n" + clientPhone + "\n\n" + "What will you be using it for?\n" + clientUsage + "\n\n" +
                "Usage Explanation:\n" + clientUsageExplanation + "\n\n" +
                 "Number of Guests:\n" + clientNumGuest + "\n\n" +
                   "Will any Non-NYU individuals be present?\n" + clientNonNYUGuest + "\n\n" +
                     "If yes, list the expected non-NYU guest here:\n" + clientNonNYUGuestList + "\n\n" +
                       "I agree to the following <a target=\"_blank\" href=\"https://docs.google.com/document/d/1TE0SI2IRChCOT5g7Xs4mDNVhlQz1KEl7pTBZFFwiJq0/edit\">Terms of Use</a>.:\n" +
                         clientAgreedToU;
      
  var event = calendar.createEvent(title, startTime, endTime, {description:description,location:location});
  Logger.log("Calendar event successfully created for client " + clientName);
  return event.getId();
}

function formatDate(date) {
  var monthNames = [
    "Jan", "Feb", "Mar",
    "Apr", "May", "Jun", "Jul",
    "Aug", "Sep", "Oct",
    "Nov", "Dec"
  ];

  var day = date.getDate();
  var monthIndex = date.getMonth();
  var year = date.getFullYear();

  return monthNames[monthIndex] + " " + day + ", " + year;
}

function pushEmails(data)
{
  var startTime = new Date(data[0][2]);
  var duration = data[0][3];
  var endTime = new Date(startTime.getTime() + data[0][3] * 60 * 1000);
  var clientName = data[0][4];
  var clientEmail = data[0][5];
  var clientNetID = data[0][6];
  var clientPhone = data[0][7];
  var clientNumGuest = data[0][8];
  var clientConsoleOption = data[0][9];
  var clientUsage = data[0][10];
  var clientUsageExplanation = data[0][11];
  var clientNonNYUGuest = data[0][12];
  var clientNonNYUGuestList = data[0][13];
  var clientAgreedToU = data[0][14];
  
  var weekday = new Array(7);
  weekday[0] = "Sun";
  weekday[1] = "Mon";
  weekday[2] = "Tue";
  weekday[3] = "Wed";
  weekday[4] = "Thu";
  weekday[5] = "Fri";
  weekday[6] = "Sat";
  
  if(clientUsageExplanation === "")
  {
    clientUsageExplanation = "N/A"; 
  }
  if(clientNonNYUGuestList === "")
  {
    clientNonNYUGuestList = "N/A"; 
  }
  
  var title = "SCHEDULED: " + clientName + " - " + clientConsoleOption + " - " + duration + " Minutes";
  var message = "<table cellpadding=\"0\" cellspacing=\"0\" align=\"center\" bgcolor=\"#ffffff\" style=\"margin-top: 15px; margin-bottom: 25px; font-family: Tahoma,Apple SD Gothic Neo,Geneva,sans-serif;\"><tbody>" +
    "<tr><td style=\"padding: 0\"><!--[if gte mso 9]><table id=\"tableForOutlook\" style=\"width:600px; height:0px;font-family:Tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 12px;\"  cellpadding=\"0\" cellspacing=\"0\">" +
    "<tr><td><![endif]--><table style=\"border: 1px solid #bbc2d0; font-family: Tahoma,Apple SD Gothic Neo,Geneva,sans-serif;font-size: 12px; width: 100%; max-width: 598px;\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\"><tbody>" +
    "<tr><td valign=\"top\" align=\"left\" width=\"100%=\" style=\"padding: 0;\"><table width=\"100%\" cellpadding=\"0\"cellspacing=\"0\" border=\"0\" valign=\"top\"><tbody><tr><td height=\"8\" width=\"100%\" style=\"" +
    "background-color: #038900;height: 8px; width: 100%; font-size: 0=px; width: 100%;\" align=\"left\" valign=\"top\"></td></tr><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" align=\"center\"" +
    "contenteditable=\"false\" height=\"29\"></td></tr></tbody></table></td></tr><tr><td valign=\"top\" align=\"left\"><table class=\"wrapper\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">" +
    "<tbody><tr><td width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" align=\"top\" valign=\"left\"><tbody><tr><td style=\"width:20px !important;\"" +
    "width=\"20\" valign=\"top\" align=\"left\"><table width=\"100%\"cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr><td style=\"font-size: 1px; line-height: 1px\"" + 
    "valign=\"top\" align=\"center\" contenteditable=\"false\"></td></tr></tbody></table></td><td valign=\"top\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr><td width=\"100%\"" +
    "valign=\"top\"align=\"center\" style=\"font-family: Tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 22px; line-height: 28px;color: #038900; font-weight: normal; width: 100%\"><inline id=\"HeaderText\">" +
    "The walk-in booking with <span style=\"word-wrap:break-word;\">" + clientName + "</span> is confirmed</inline></td></tr><tr><td style=\"font-size: 1px; line-height: 1px\"" +
    "valign=\"top\" align=\"left\" contenteditable=\"false\" height=\"27\"></td></tr><tr><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\"" +
    "cellpadding=\"0\" cellspacing=\"0\" border=\"0\"class=\"flexible_mobile\"><tbody><tr><td style=\"font-size: 1px;line-height: 1px\" valign=\"top\" align=\"left\" contenteditable=\"false\"" +
    "height=\"3\"></td></tr><tr><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: Tahoma,Apple SD Gothic Neo,Geneva,sans-serif;font-size: 14px;line-height: 19px; color: #333; font-weight:" +
    "normal;\"><inline id=\"TopText\">The walk-in booking with <span style=\"word-wrap:break-word;\">" + clientName + "</span> (<span>" + clientEmail + "</span>) is confirmed. Please see below for more information.</inline>" +
    "</td></tr><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" align=\"left\" contenteditable=\"false\" height=\"15\"></td></tr></tbody></table></td></tr></tbody></table></td>" +
    "<td style=\"width:20px !important;\" width=\"20\" valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr><td style=\"font-size: 1px;" +
    "line-height: 1px\" valign=\"top\" align=\"center\" contenteditable=\"false\"></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr><td valign=\"top\" align=\"left\" style=\"font-size: 1px;" +
    "line-height: 1px;\"contenteditable=\"false\" height=\"\"><table width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" bgcolor=\"\"><tbody><tr><td style=\"height: 22px; width: 45%; background: #fff;\"" +
    "height=\"22\" align=\"left\">&nbsp;</td><td width=\"66\" style=\"width: 66px; padding: 0\" rowspan=\"3\" valign=\"top\" background=\"transparant\" align=\"center\"><a href=\"#\"" +
    "style=\"text-decoration: none;\"><img width=\"66\" height=\"50\" style=\"border: 0; display: block; position: relative; z-index: 3; border: 1px solid #fff; width: 66px; height: 50px\" rel=\"border: 0; display: block;" +
    "position: relative; z-index: 3; border: 1px solid #fff; width: 66px; height: 50px\" src=\"http://static.scheduleonce.com/Images/email/bookingDetailsIcon.jpg\" alt=\"\"></a></td><td style=\"height: 22px; width: 45%;" +
    "background: #fff;\" height=\"22\" align=\"right\">&nbsp;</td></tr><tr><td style=\"background-color: #cbcbcb; height: 2px; width: 45%; font-size: 0px\" height==\"2\" align=\"left\">&nbsp;</td><td " +
    "style=\"background-color: #cbcbcb; height: 2px; width: 45%; font-size: 0px\" height=\"2\" align=\"right\">&nbsp;</td></tr><tr><td style=\"background-color: #ffffff; height: 20px; width: 45%; font-size: 0px;" +
    "line-height: 0\" height=\"20\" align=\"left\">&nbsp;</td><td style=\"background-color: #ffffff; height: 20px; width: 45%; font-size: 0px; line-height: 0\" height=\"20\" align=\"right\">&nbsp;" +
    "</td></tr></tbody></table></td></tr><tr><td width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr><td style=\"width:20px !important;\"" +
    "width=\"20\" valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\"" +
    "align=\"center\" contenteditable=\"false\"></td></tr></tbody></table></td><td valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr>" +
    "<td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" align=\"left\" contenteditable=\"false\" height=\"13\"></td></tr><tr><td valign=\"top\" align=\"left\"" +
    "style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 20px; line-height: 24px; color: #333; font-weight: normal;\">Booking details</td></tr><tr><td style=\"font-size: 1px; line-height: 1px\"" +
    "valign=\"top\" align=\"left\" contenteditable=\"false\" height=\"15\"></td></tr><tr><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\"" +
    "cellspacing=\"0\" border=\"0\"><tbody><tr><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px;font-weight: bold; line-height: 21px;" +
    "color: #333333;\">Duration</td></tr><tr><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px;line-height: 21px; color: #333333;" +
    "font-weight: normal;\"><span style=\"word-wrap:break-word;\">" + clientConsoleOption + " - " + duration + " Minutes" + "</tr><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" align=\"left\"" +
    "contenteditable=\"false\" height=\"21\"></td></tr></tbody></table></td></tr><tr><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\"" +
    "cellspacing=\"0\" border=\"0\"><tbody><tr><td width =\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px; font-weight: bold; line-height: 21px;" +
    "color: #333333;\">NYU NetID</td></tr><tr><td width =\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px; line-height: 21px; color: #333333;" +
    "font-weight: normal;\">" + clientNetID + "</td></tr><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" contenteditable=\"false\" height=\"21\" align=\"left\"></td></tr></tbody></table>" +
    "</td></tr><tr id=\"D-11_header\"><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody>" +
    "<tr id=\"D-11_title\"><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px;=font-weight: bold; line-height: 21px;" +
    "color: #333333;\">Phone Number</td></tr><tr id=\"D-11\"><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px;line-height: 21px;" +
    "color: #333333; font-weight: normal;\"><span style=\"word-wrap:break-word;\">" + clientPhone + "</span></td></tr><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" align=\"left\"" +
    "contenteditable=\"false\" height=\"21\"></td></tr></tbody></table></td></tr><tr id=\"D-2_header\"><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\"" +
    "cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr id=\"D-2_title\"><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif;" +
    "font-size: 14px;font-weight: bold; line-height: 21px; color: #333333;\"><inline id=\"OwnerTimeText\">Your Reservation Time</inline></td></tr><tr id=\"D-2\"><td width=\"100%\" valign=\"top\" align=\"left\"" +
    "style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px;line-height: 21px; color: #333333; font-weight: normal;\"><span style=\"word-wrap:break-word;\">" +
    weekday[startTime.getDay()] + ", " + formatDate(startTime) + "</span></td></tr><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" align=\"left\" contenteditable=\"false\" height=\"21\"></td>" +
    "</tr></tbody></table></td></tr><tr id=\"D-1_header\"><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">" +
    "<tbody><tr id=\"D-1_title\"><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px;font-weight: bold; line-height: 21px; color: #333333;\">" +
    "<inline id=\"CustomerTimeText\">Reservation Duration</inline></td></tr><tr id=\"D-1\"><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif;" +
    "font-size: 14px;line-height: 21px; color: #333333; font-weight: normal;\"><span style=\"word-wrap:break-word;\">" + duration + " Minutes</span></td></tr><tr><td style=\"font-size: 1px; line-height: 1px\"" +
    "valign=\"top\" align=\"left\" contenteditable=\"false\" height=\"21\"></td></tr></tbody></table></td></tr><tr id=\"D-9_header\"><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\">" +
    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr id=\"D-9_title\"><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva," +
    "sans-serif; font-size: 14px; font-weight: bold; line-height: 21px; color: #333333;\">VR Play Area</td></tr><tr id=\"D-9\"><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family:" +
    "tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px; line-height: 21px; color: #333333; font-weight: normal;\"><span style=\"word-wrap:break-word;\">539-541 LaGuardia Pl. New York, NY 10012" +
    "(<a href=\"https://maps.google.com/?q=539-541 LaGuardia Pl. New York, NY 10012\" target=\"_blank\">map</a>)</span></td></tr><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" align=\"left\"" +
    "contenteditable=\"false\" height=\"21\"></td></tr></tbody></table></td></tr><tr id=\"D-18_header\"><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\"" +
    "cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px;font-weight:" +
    "bold; line-height: 21px; color: #333333;\">Booking Option</td></tr><tr id=\"D-18\"><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px;" +
    "line-height: 21px; color: #333333; font-weight: normal;\"><span style=\"word-wrap:break-word;\">Walk-in Booking</span></td></tr><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" align=\"left\"" +
    "contenteditable=\"false\" height=\"14\"></td></tr></tbody></table></td></tr></tbody></table></td><td style=\"width:20px !important;\" width=\"20\" valign=\"top\" align=\"left\"><table width=\"100%\"" +
    "cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" align=\"center\"></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr>" +
    "<td valign=\"top\" align=\"left\" style=\"font-size: 1px; line-height: 1px;\" contenteditable=\"false\" height=\"\"><table width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" bgcolor=\"\">" +
    "<tbody><tr><td style=\"height: 22px; width: 45%; background: #fff;\" height=\"22\" align=\"left\">&nbsp;</td><td width=\"66\" style=\"width: 66px; padding: 0 !important\" rowspan=\"3\" valign=\"top\"" +
    "background=\"transparant\" align=\"center\"><a href=\"#\" style=\"text-decoration: none;\"><img style=\"border: 0; display: block; position: relative; z-index: 3; border: 1px solid #fff; width: 66px;" +
    "height: 50px\" rel=\"border: 0; display: block;position: relative; z-index: 3; border: 1px solid #fff; width: 66px; height: 50px\" src=\"http://static.scheduleonce.com/Images/email/informationIcon.jpg\" alt=\"\"" + 
    "width=\"66\" height=\"50\"></a></td><td style=\"height: 22px; width: 45%; background: #fff;\" height=\"22\" align=\"right\">&nbsp;</td></tr><tr><td style=\"background-color: #cbcbcb; height: 2px; width: 45%;" +
    "font-size: 0px\" height=\"2\" align=\"left\">&nbsp</td><td style=\"background-color: #cbcbcb; height: 2px; width: 45%; font-size: 0px\" height=\"2\" align=\"right\">&nbsp;</td></tr><tr>" +
    "<td style=\"background-color: #ffffff; height: 20px; width: 45%; font-size: 0px; line-height: 0\" height=\"20\" align=\"left\">&nbsp;</td><td style=\"background-color: #ffffff; height: 20px; width: 45%; font-size: 0px;" +
    "line-height: 0\" height=\"20\" align=\"right\">&nbsp;</td></tr></tbody></table></td></tr><tr><td width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\"" +
    "border=\"0\"><tbody><tr><td style=\"width:20px !important;\" width=\"20\" valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr>" +
    "<td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" align=\"center\" contenteditable=\"false\"></td></tr></tbody></table></td><td valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\"" +
    "cellspacing=\"0\" border=\"0\"><tbody><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" align=\"left\" contenteditable=\"false\" height=\"13\"></td></tr><tr><td valign=\"top\" align=\"left\"" +
    "style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 20px; line-height: 24px; color: #333; font-weight: normal;\">Additional details</td></tr><tr><td style=\"font-size: 1px; line-height: 1px\"" +
    "valign=\"top\" align=\"left\" contenteditable=\"false\" height=\"15\"></td></tr><tr><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\"" +
    "cellspacing=\"0\" border=\"0\"><tbody><tr><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px;font-weight: bold; line-height: 21px;" +
    "color: #333333;\"><inline id=\"S-2\" style=\"word-wrap:break-word\">Client name</inline></td></tr><tr><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva," +
    "sans-serif; font-size: 14px;line-height: 21px; color: #333333; font-weight: normal;\"><span style=\"word-wrap:break-word;\">" + clientName + "</span></td></tr><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\"" +
    "align=\"left\" contenteditable=\"false\" height=\"21\"></td></tr></tbody></table></td></tr><tr id=\"D-21_header\"><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\"" +
    "cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px;font-weight: bold;" +
    "line-height: 21px; color: #333333;\"><inline id=\"S-5\" style=\"word-wrap:break-word\">Number of Guests</inline></td></tr><tr id=\"D-21\"><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family:" +
    "tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px;line-height: 21px; color: #333333; font-weight: normal;\"><span style=\"word-wrap:break-word;\">" + clientNumGuest + "</span></td></tr><tr><td style=\"font-size: 1px;" +
    "line-height: 1px\" valign=\"top\" align=\"left\" contenteditable=\"false\" height=\"21\"></td></tr></tbody></table></td></tr><tr id=\"D-25_header\"><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\"" +
    "align=\"left\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family: tahoma,Apple SD Gothic Neo,Geneva," +
    "sans-serif; font-size: 14px;font-weight: bold; line-height: 21px; color: #333333;\">Usage Reason</td></tr><tr id=\"D-25\"><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family:" +
    "tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size: 14px;line-height: 21px; color: #333333; font-weight: normal;\"><span style=\"word-wrap:break-word;\">" + clientUsage + "</span></td></tr><tr><td style=\"font-size: " +
    "1px; line-height: 1px\" valign=\"top\" align=\"left\" contenteditable=\"false\" height=\"21\"></td></tr></tbody></table></td></tr><tr><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\">" +
    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody></tbody></table></td></tr><tr><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\">" +
    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td width=\"100%\" valign=\"top\" align=\"left\" style=\"font-family:tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size:" +
    "14px; font-weight:bold; line-height:21px; color:#333333;\">What will you be using it for?</td></tr><tr><td width=\"100%\" valign =\"top\" align = \"left\" style=\"font-family:tahoma,Apple SD Gothic Neo,Geneva,sans-serif;" +
    "font-size:14px; line-height:21px; color:#333333; font-weight:normal;\">" + clientUsageExplanation + "</td></tr></table></td></tr><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" height=\"21\"" +
    "align=\"left\"></td></tr><tr><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td width=\"100%\"" +
    "valign=\"top\" align=\"left\" style=\"font-family:tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size:14px; font-weight:bold; line-height:21px; color:#333333;\">Will any Non-NYU individuals be present?</td></tr>" +
    "<tr><td width=\"100%\" valign =\"top\" align = \"left\" style=\"font-family:tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size:14px; line-height:21px; color:#333333; font-weight:normal;\">" + clientNonNYUGuest + 
    "</td></tr></table></td></tr><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" height=\"21\" align=\"left\"></td></tr><tr><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\">" +
    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td width=\"100%\" valign=\"top\"align=\"left\" style=\"font-family:tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size:14px;" +
    "font-weight:bold; line-height:21px; color:#333333;\">If yes, list the expected non-NYU guest here</td></tr><tr><td width=\"100%\" valign =\"top\" align = \"left\" style=\"font-family:tahoma,Apple SD Gothic Neo,Geneva," +
    "sans-serif; font-size:14px; line-height:21px; color:#333333; font-weight:normal;\">" + clientNonNYUGuestList + "</td></tr></table></td></tr><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" height=\"21\"" +
    "align=\"left\"></td></tr><tr><td class=\"flexible_mobile\" width=\"100%\" valign=\"top\" align=\"left\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td width=\"100%\"" +
    "valign=\"top\" align=\"left\" style=\"font-family:tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size:14px; font-weight:bold; line-height:21px; color:#333333;\">I agree to the following&nbsp;" + 
    "<a target=_blank href=https://docs.google.com/document/d/1TE0SI2IRChCOT5g7Xs4mDNVhlQz1KEl7pTBZFFwiJq0/edit>Terms of Use</a>.</td></tr><tr><td width=\"100%\" valign =\"top\" align = \"left\"" +
    "style=\"font-family:tahoma,Apple SD Gothic Neo,Geneva,sans-serif; font-size:14px; line-height:21px; color:#333333; font-weight:normal;\">" + clientAgreedToU + "</td></tr></table></td></tr><tr><td style=\"font-size: 1px;" +
    "line-height: 1px\" valign=\"top\" height=\"21\" align=\"left\"></td></tr></tbody></table></td><td style=\"width:20px !important;\" width=\"20\" valign=\"top\" align=\"left\"><table width=\"100%\"" +
    "cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tbody><tr><td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" align=\"center\"></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr>" +
    "<td style=\"font-size: 1px; line-height: 1px\" valign=\"top\" align=\"center\" contenteditable=\"false\" height=\"15\"></td></tr></tbody></table></td></tr></tbody></table><!--[if gte mso 9]></td></tr>" +
    "</table><![endif]--></td></tr></tbody></table>";
  
  var sender = GmailApp.getAliases()[0] + "";
  
  GmailApp.sendEmail(
    clientEmail,         
    title,               
    'test', {           
      from: sender,
      name: 'LaGuardia Co-op',
      replyTo: 'student.tech.centers@nyu.edu',
      htmlBody: message
    }
  ); 
  Logger.log("Email notification successfully created for client " + clientName);
  
}

function extractClientName(data) {
  try{
    Logger.log(data);
    Logger.log(data.match(/Attendee:\n*(.*)[(](.*?)[)]/g));
    var name = data.match(/Attendee:\n*(.*)[(](.*?)[)]/g)[1];
    return name;
  }
  catch(err)
  {
    Logger.log(err);
    return "";
  }
}

function extractClientEmail(data) {
  try{
    var email = data.match(/Attendee:\n*(.*)[(](.*?)[)]/g)[2];
    return email;
  }
  catch(err)
  {
    Logger.log(err);
    return "";
  }
}

function extractClientPhone(data) {
  try{
    var phone = data.match(/Phone:\n*(.*)/)[2];
    return phone;
  }
  catch(err)
  {
    return "";
  }
}

function extractClientNetID(data) {
  try{
    var netID = data.match(/NYU\sNet\sID:\n*([A-z0-9]*)/)[1];
    return netID;
  }
  catch(err)
  {
    return "";
  }
}

function extractClientNumGuest(data) {
  try{
    var guests = data.match(/Expected\snumber\sof\sattendees:\n*([0-9]*)/)[1];
    return guests;
  }
  catch(err)
  {
    return "";
  }
}

function extractClientConsoleOption(data) {
  try{
    var console = data.match(/Console:\n*(.*)/)[1];
    return console;
  }
  catch(err)
  {
    return "";
  }
}

function extractClientUsage(data) {
  try{
    var usage = data.match(/Usage:\n*(.*)/)[1];
    return usage;
  }
  catch(err)
  {
    return "";
  }
}

function extractClientUsageExplanation(data) {
  try{
    var reason = data.match(/If Development\/Digital Sculpting, briefly explain project:\n*(.*)/)[1];
    return reason;
  }
  catch(err)
  {
    return "";
  }
}

function extractClientNonNYUGuest(data) {
  try{
    var nonnyu = data.match(/Will any Non-NYU individuals be present during the session\?:\n*(.*)/)[1];
    return nonnyu;
  }
  catch(err)
  {
    return "";
  }
}

function extractClientNonNYUGuestList(data) {
  try{
    var list = data.match(/If yes, list the expected non-NYU guest here:\n*(.*)/)[1];
    return list;
  }
  catch(err)
  {
    return "";
  }
}

function extractClientAgreedToU(data) {
  try{
    var tou = data.match(/I agree to the following <a target=_blank href=https:\/\/docs\.google\.com\/document\/d\/1TE0SI2IRChCOT5g7Xs4mDNVhlQz1KEl7pTBZFFwiJq0\/edit>Terms of Use<\/a>.:\n*(.*)/)[1];
    return tou;
  }
  catch(err)
  {
    return "";
  }
}