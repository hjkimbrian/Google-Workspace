#!/bin/bash
gam=$HOME/bin/gamadv-xtd3

#Purpose of the script
#To identify recurring events that are internal with attendees of greater than 1
#This will delete all recurrences of events, whether the specific recurrences are in the past or in the future.

#READ ME FIRST
#This requires running GAMADV-XTD3 5.25.17 or higher
#https://github.com/taers232c/GAMADV-XTD3/releases/tag/v5.25.17
#To update: Copy and paste following: bash <(curl -s -S -L https://git.io/fhZWP) -l
#https://github.com/taers232c/GAMADV-XTD3/wiki/CSV-Output-Filtering 
#https://github.com/taers232c/GAMADV-XTD3/wiki/Users-Calendars-Events


#Step 1 - list recurring events in each user's primary calendar
# this creates a CSV file named eventstoDelete.csv

$gam config csv_output_row_filter "'^attendees$:count>1','recurrence:count>=1','attendees.*email:all:regex:(^$)|(.+@domain.com)'" redirect csv ./eventstoDelete.csv multiprocess all users print events primary

#Step 2 - examine the CSV file and make sure that the events make sense
#Following command will delete all recurrences of events. PROCEED WITH CAUTION. Uncomment #doit if you want to delete the events. 

$gam csv eventstoDelete.csv gam user "~primaryEmail" delete events "~calendarId"  id "~id" sendnotifications false #doit
