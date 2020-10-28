function deleteCalendarEvents() {
  //All credit goes to @jbastean on MacAdmins
  //https://macadmins.slack.com/archives/C18SZ6S07/p1566334369190700
  // ****************************
  // UPDATE THESE VARIABLES
  // ****************************
  
  // Calendar ID - hover over calendar name on left, click 3 dots -> "Settings & Sharing" -> "Calendar Id" (near bottom)
  // Should be in format similar to 'domain.edu_abcd1234foobar@group.calendar.google.com
  // Resource calendars will be similar, but have '@resource.calendar.google.com'
    
  var calendarArray = [  // Add array of calendarIDs here
    'exampledomain.edu_1bcd123@group.calendar.google.com',
    'exampledomain.edu_123foobar987@group.calendar.google.com'
  ]
  
  // Date format: 'February 17, 2016 13:00:00 -0500'
  
  var startDate = new Date('August 18, 2019 00:00:00 -0500')
  var endDate = new Date('December 30, 2019 00:00:00 -0500')
  
  // *****************************
  // NO CHANGES BEYOND THIS POINT
  // *****************************
  
  for (var i = 0; i < calendarArray.length; i++) {
    // Get array of Calendar events
    var events = CalendarApp.getCalendarById(calendarArray[i]).getEvents(startDate, endDate)
    
    // Loop through array, deleting each event
    for (var j = 0; j < events.length; j++) {
      // Logger.log(events[i].getTitle());  // Uncomment to log titles of each event. Useful to make sure you targeted the correct calendar before deleting. 
      // Comment out next line if doing this.
      events[j].deleteEvent();
      Utilities.sleep(150)  // Adjust as needed to not hit rate limiting. Usually 125 works, but I ran into rate limiting occasionally, so make sure to babysit if you reduce it.
    }
  }
  
  
}
