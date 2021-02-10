/*
Purpose of the script: sync calendar events from one calendar to the other & avoid creating duplicates
Adapted from: https://medium.com/@willroman/auto-block-time-on-your-work-google-calendar-for-your-personal-events-2a752ae91dab
Warning: does contain a bug, use with caution. 
*/
function OtherToWork() {

  var id=""; // CHANGE - id of the secondary calendar to pull events from

  var today=new Date();
  var enddate=new Date();
  enddate.setDate(today.getDate()+30); // how many days in advance to monitor and block off time
  
  var secondaryCal=CalendarApp.getCalendarById(id);
  var secondaryEvents=secondaryCal.getEvents(today,enddate);
  
  var primaryCal=CalendarApp.getDefaultCalendar();
  var primaryEvents=primaryCal.getEvents(today,enddate); // all primary calendar events
  
  var primaryEventTitle="Busy - Other"; // update this to the text you'd like to appear in the new events created in primary calendar
  var duplicateEventTitle="Busy - Work/Personal"; // add any calendar event titles that you do not want copied back
  
  var stat=1;
  var evi, existingEvent; 
  var primaryEventsFiltered = []; // to contain primary calendar events that were previously created from secondary calendar
  var primaryEventsUpdated = []; // to contain primary calendar events that were updated from secondary calendar
  var primaryEventsCreated = []; // to contain primary calendar events that were created from secondary calendar
  var primaryEventsDeleted = []; // to contain primary calendar events previously created that have been deleted from secondary calendar

  Logger.log('Number of primaryEvents: ' + primaryEvents.length);  
  Logger.log('Number of secondaryEvents: ' + secondaryEvents.length);
  
  // create filtered list of existing primary calendar events that were previously created from the secondary calendar
  for (pev in primaryEvents)
  {
    var pEvent = primaryEvents[pev];
    if (pEvent.getTitle() == primaryEventTitle)
    { primaryEventsFiltered.push(pEvent); }
  }
  
  // process all events in secondary calendar
  for (sev in secondaryEvents)
  {
    stat=1;
    evi=secondaryEvents[sev];
    
 
    // skip duplicates
    for (existingEvent in primaryEvents)
      {
        var pEvent = primaryEvents[existingEvent];
        var secondaryTitle = evi.getTitle();
        var secondaryDesc = evi.getDescription();
        if ((pEvent.getStartTime().getTime()==evi.getStartTime().getTime()) && (pEvent.getEndTime().getTime()==evi.getEndTime().getTime() && secondaryTitle!==duplicateEventTitle))
        {
          stat=0;
          break; 
          // pEvent.setTitle(primaryEventTitle);
          // pEvent.setDescription(secondaryTitle + '\n\n' + secondaryDesc);
          // event.setDescription(evi.getTitle() + '\n\n' + evi.getDescription());
          // pEvent.setVisibility(CalendarApp.Visibility.PRIVATE); // set blocked time as private appointments in work calendar
          // primaryEventsUpdated.push(pEvent.getId());
          //Logger.log('PRIMARY EVENT UPDATED'
          //           + '\nprimaryId: ' + pEvent.getId() + ' \nprimaryTitle: ' + pEvent.getTitle() + ' \nprimaryDesc: ' + pEvent.getDescription());
        } 
      }
    
    // if the secondary event has already been blocked in the primary calendar, update it
    for (existingEvent in primaryEventsFiltered)
      {
        var pEvent = primaryEventsFiltered[existingEvent];
        var secondaryTitle = evi.getTitle();
        var secondaryDesc = evi.getDescription();
        if ((pEvent.getStartTime().getTime()==evi.getStartTime().getTime()) && (pEvent.getEndTime().getTime()==evi.getEndTime().getTime()))
        {
          stat=0;
          pEvent.setTitle(primaryEventTitle);
          pEvent.setDescription(secondaryTitle + '\n\n' + secondaryDesc);
          // event.setDescription(evi.getTitle() + '\n\n' + evi.getDescription());
          pEvent.setVisibility(CalendarApp.Visibility.PRIVATE); // set blocked time as private appointments in work calendar
          primaryEventsUpdated.push(pEvent.getId());
          Logger.log('PRIMARY EVENT UPDATED'
                     + '\nprimaryId: ' + pEvent.getId() + ' \nprimaryTitle: ' + pEvent.getTitle() + ' \nprimaryDesc: ' + pEvent.getDescription());
        } 
      }

    if (stat==0) continue;    
    
    var d = evi.getStartTime();
    var n = d.getDay();

    if (evi.isAllDayEvent())
    {
      continue; // Do nothing if the event is an all-day or multi-day event. This script only syncs hour-based events
    }
    else if (n==1 || n==2 || n==3 || n==4 || n==5 && secondaryTitle!==duplicateEventTitle) // skip weekends and duplicates. Delete this if you want to include weekends
    // if the secondary event does not exist in the primary calendar, create it
    {
      var newEvent = primaryCal.createEvent(primaryEventTitle,evi.getStartTime(),evi.getEndTime()); // change the Booked text to whatever you would like your merged event titles to be
      // alternative version below that copies the exact secondary event information into the primary calendar event
      // var newEvent = primaryCal.createEvent(evi.getTitle(),evi.getStartTime(),evi.getEndTime(), {location: evi.getLocation(), description: evi.getDescription()});  
      // newEvent.setDescription(evi.getTitle() + '\n\n' + evi.getDescription());
      newEvent.setVisibility(CalendarApp.Visibility.PRIVATE); // set blocked time as private appointments in work calendar
      newEvent.removeAllReminders(); // so you don't get double notifications. Delete this if you want to keep the default reminders for your newly created primary calendar events
      primaryEventsCreated.push(newEvent.getId());
      Logger.log('PRIMARY EVENT CREATED'
                 + '\nprimaryId: ' + newEvent.getId() + '\nprimaryTitle: ' + newEvent.getTitle() + '\nprimaryDesc ' + newEvent.getDescription() + '\n');
    }
  }

  // if a primary event previously created no longer exists in the secondary calendar, delete it
  for (pev in primaryEventsFiltered)
  {
    var pevIsUpdatedIndex = primaryEventsUpdated.indexOf(primaryEventsFiltered[pev].getId());
    if (pevIsUpdatedIndex == -1)
    { 
      var pevIdToDelete = primaryEventsFiltered[pev].getId();
      Logger.log(pevIdToDelete + ' deleted');
      primaryEventsDeleted.push(pevIdToDelete);
      primaryEventsFiltered[pev].deleteEvent();
    }
  }  

  Logger.log('Primary events previously created: ' + primaryEventsFiltered.length);
  Logger.log('Primary events updated: ' + primaryEventsUpdated.length);
  Logger.log('Primary events deleted: ' + primaryEventsDeleted.length);
  Logger.log('Primary events created: ' + primaryEventsCreated.length);

}  
