# Excel Tournament Assistant
When procrastinating examn preperations, I started working on an automatic
Excel spreadsheet that utilizes macros to dynamically generate tournament
brackets for when me and my friends get together in our seasonal smash tournaments.


## Running the sheet
Running the sheet requires you to make sure that macros are enabled.
Please have a look at these to convince yourself that there is no ill intent 
in the VBA code.

## Usage
Upon opening the xlsm file, a column of participants is filled in. 
Once the list is complete, run the group stage macro by pressing the **generate**
button in the first (group stage) sheet. This generates a round robin system with
a points table as well as a live standings table.

### Setting preferences
In the _Preferences_ sheet, you can customize the look of the workbook as well as setting 
several parameters for the assistant. 

### Points table
The points table keeps track of all matches that are played between the players.

### Matches
The matches are generated such that each player will face the others once during the 
group stage. For odd numbers of participants, one player will face _[NONE]_ each round. This is just a dummy match and can be disregarded. 
The workbook will warn you about matches that have 
