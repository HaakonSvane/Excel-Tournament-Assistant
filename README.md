# Excel Tournament Assistant
When procrastinating examn preperations, I started working on an automatic
Excel spreadsheet that utilizes macros to dynamically generate tournament
brackets for when me and my friends get together in our seasonal Super Smash Brothers tournaments.


## Running the sheet
Running the sheet requires you to make sure that macros are enabled.
Please have a look at these to convince yourself that there is no ill intent 
in the VBA code.

## Usage
### Setting preferences
In the _Preferences_ sheet, you can customize the look of the workbook as well $
several parameters for the assistant. These include:
 * Best of: Seperate values can be set for groupstage, tiebreaker games and the$
 * Winner bracket advantage: If enabled, the finalist from the lower bracket fi$
 * Participants limit: Dumb parameter that I should get rid of. This is used to$

### Group stage sheet
Upon opening the .xlsm file, a column of participants is filled in. 
Once the list is complete, run the group stage macro by pressing the **generate**
button in the first (group stage) sheet. This generates a round robin system with
a points table as well as a live standings table. 

#### Points table
The points table keeps track of all matches that are played between the players.

#### Matchups
The matches are generated such that each player will face the others once during the 
group stage. For odd numbers of participants, one player will face _[NONE]_ each round. This is just a dummy match and can be disregarded. 
The workbook will warn you about matches that have impossibles scores, for example negative points or points that exceed the *best of* value.

#### Standings
The standings panel is a live view of the scores of each participant.
It keeps track of total points, number of matches played and number of victories. One point is given for each victory in a set so that in a BO3, the maximum number of points is two.
The column with number of games will be colored green once a player has finished all their games in the group stage.


