/* Delete Cancelled Events
 * By Eric Lin
 *
 * This script deletes events that are prefixed with "CANCELED"
 * Hopefully this script doesn't delete false positives
 * Currently, no false positives has been detected.
 */

function deleteCancelledEvents() {
  var calendarID = ['nyu.edu_6e79752d6c6167636f6f702d623037@resource.calendar.google.com',
                   'nyu.edu_6e79752d6c6167636f6f702d623039@resource.calendar.google.com',
                   'nyu.edu_6e79752d6c6167636f6f702d623032@resource.calendar.google.com',
                   'nyu.edu_6e79752d6c6167636f6f702d313039@resource.calendar.google.com',
                   'nyu.edu_6e79752d6c6167636f6f702d313037@resource.calendar.google.com',
                   'nyu.edu_6e79752d6c6167636f6f702d313034@resource.calendar.google.com',
                   'nyu.edu_pshid3c02pqduechcf9a3hjbt8@group.calendar.google.com'];
  
  var now = new Date();
  var oneWeekAgo = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));
  var oneWeekFromNow = new Date(now.getTime() + (7 * 24 * 60 * 60 * 1000));
  var deletedEvents = 0;
  for (var i = 0; i < calendarID.length; i++)
  {
    var calendar = CalendarApp.getCalendarById(calendarID[i]);
    var events = calendar.getEvents(oneWeekAgo, oneWeekFromNow,{search: 'CANCELED'});
    Logger.log('# canceled in calendar (' + calendar.getName() + '): ' + events.length);
    for (var j = 0; j < events.length; j++)
    {
      Logger.log('About to delete ' + events[j].getTitle() + ' on ' + events[j].getStartTime());
      events[j].deleteEvent();
      deletedEvents += 1;
    }
  }
  if(deletedEvents > 0)
  {
    var emailAddress = "eric.lin@nyu.edu";
    var subject = "Delete Cancelled Events Logs - " + new Date(); 
    var message = Logger.getLog();
    MailApp.sendEmail(emailAddress, subject, message);
  }
}
