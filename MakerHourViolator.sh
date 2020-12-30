#!/bin/bash
gam=$HOME/bin/gamadv-xtd3/gam

#Purpose of the script
#To identify events that are booked during "Maker Hour" - twice a week on Tues and Thursday
#Just need to identify the organizers that are booking the meetings, no need to delete

#Set variables for maker hour
date=2020-12-15
timezone=05:00
makerhourstart=13:00:00
makerhourend=17:00:00

#list events that fall within the date and time range whose title is not "Maker Hour"
#Replace Maker Hour as needed in summary:notregex:Maker Hour. Space does not need to be escaped with a backslash.
#This create a new Google Sheet with a title "Maker Hour Violators - $date"
#Use loops and if statements in shell to loop through multiple dates and append to the same sheet as required.
#More info: https://github.com/taers232c/GAMADV-XTD3/wiki/Todrive
#More info: https://github.com/taers232c/GAMADV-XTD3/wiki/Users-Calendars-Events#display-calendar-events
#More info: https://github.com/taers232c/GAMADV-XTD3/wiki/CSV-Output-Filtering#column-row-filtering
#Make pivot tables as needed in Google Sheet to report on the number of unique Ids by Organizer.email

$gam config csv_output_row_filter "'summary:notregex:^Maker Hour$'" auto_batch_min 1 redirect csv - multiprocess todrive tdtitle "Maker Hour Violators -"${date} all users print events primary starttime ${date}T${makerhourstart}-${timezone} before ${date}T${makerhourend}-${timezone} fields summary,organizer.email
