prsnl_upload


This script doese the following:

Queries a xlsx spreadsheet
Grabs the person id, username, first name, last name, position, begin date, and end date
Compiles an xml document and called prsnlimport.xml that can be directly uploaded using contentmanager.exe in the Cerner ProgramFiles


Spreadsheet Info
The spreadsheet must be formatted in the following manner:

Column header organization is arbitrary, but spelling is important

Column Headers:

username
firstname
lastname
position

Must be spelled exactly as it is in Cerner. (It is also case/white space sensitive)


begin_date


Format: yearmonthdaytime
Example: June 11, 2020 0500 = 2020061105000000


end_date


Format: yearmonthdaytime
Cerners default of December 31, 2100 0000 = 2100123105000000


middle

(Optional Field)


person_id

(Optional Field if the user does not exist in that domain already)





Note: There is a template in this repo called prsnl_import.xlsx

Log Info


Location: root\logs

Convention: Currently not using the log functionality, but the variables are setup incase it is needed


Important Notes

The script currently applies the following org groups automatically for all users being imported:

Physician Network Services
Texas Tech Physicians
UMCP


This functionality can be modified by commenting out the org groups you do not want on the current group of users you are importing.
The "department" column header is in the template provided, but not currently being used in the script. This will hopefully be implemented later on, if possible.# My-Projects
