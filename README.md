# Excel Tournament Assistant

[![GroupStage-100%](https://img.shields.io/badge/Group%20Stage-100%25-green)]()
[![MainStage-40%](https://img.shields.io/badge/Main%20Stage-40%25-orange)]()
[![Preferences-60%](https://img.shields.io/badge/Preferences-60%25-orange)]()

When procrastinating examn preperations, I started working on an automatic
Excel spreadsheet that utilizes macros to dynamically generate tournament
brackets for when me and my friends get together in our seasonal Super Smash Brothers tournaments.


## Running the sheet
Running the sheet requires you to make sure that macros are enabled.
Please have a look at these to convince yourself that there is no ill intent 
in the VBA code. The modules in the git are are __not__ meant to justify the safety of the macros as they are exported and pushed independently of the .xlsm file.
I decided to add them to github just to track diffs and to expose them easier.

## Usage
### Setting preferences

![image](/asset/img/preferencesSheet.png)

In the _Preferences_ sheet, you can customize the look of the workbook as well as several parameters for the assistant. These include:
 * Best of: Seperate values can be set for groupstage, tiebreaker games, mainstage and the mainstage finals.  
 * Winner bracket advantage: If enabled, the finalist from the lower bracket finals will need to win two sets in order to win the tournament.
 * Participants limit: Dumb parameter that I should get rid of. This is used to set a limit to the search of participant names in the initial column as well as limiting the area required to clear each time any macros run.

### Group stage sheet

![image](/asset/img/emptyGroupStage.png)

Upon opening the .xlsm file, a column of participants is filled in. 
Once the list is complete, run the group stage macro by pressing the **generate**
button in the first (group stage) sheet. This generates a round robin system with
a points table as well as a live standings table.

![image](/asset/img/generatedGroupStage.png)

When all matches have been played for a player, the name of the player is highligted in the Standings table. All games have registered and are valid once all the player names in this table are highligted.
When this is done, head into the MainStage sheet to proceed.

![image](/asset/img/filledGroupStage.png)

#### Points table
The points table keeps track of all matches that are played between the players in a condensed format. Players in the left column have their points on the left of the matchup.

#### Matchups
The matches are generated such that each player will face the others once during the 
group stage. For odd numbers of participants, one player will face _[NONE]_ each round. This is a dummy match and can be disregarded. 
The workbook will warn you about matches that have forbidden scores, for example negative points or points that exceed the *best of* value.

#### Standings
The standings pane is a live view of the scores of each participant.
It keeps track of total points, number of matches played and number of victories. One point is given for each victory in a set so that in a BO3, the maximum number of points is two.
The column with number of games will be colored green once a player has finished all their games in the group stage.

### Main stage sheet

![image](/asset/img/generatedMainStage.png)

__The main stage sheet is still under development!__

#### Adjusted standings
The adjusted standings determines the outcome for each cluster of players from the group stage with equal scores. This is a more sophisticated version of the standings table in the group stage sheet that
accounts for tiebreaker games and extra points given to players with the same score (clustered players).
There are two possible ways of handling a cluster:
1. The cluster consists of __two__ players. The following rules then apply:
	* If the players are to meet in their first match of the elimination phase, their names are highlighted in green and they do __not__ need to play a tiebreaker game.
	* If the previous condition is not met, the players names will be highlighed in red. Further, a tiebreaker game will be set up automatically to the right of the adjusted standings table.
2. The cluster consists of __three or more__ players. The following rules then apply:
	1. Each player has _extra points_ added to their score for each set they have won in the group stage.
	2. Each player has _extra points_ added to their score for each set they have won against the other players in the cluster
	3. If some of the scores are still equal, the standings of the cluster will be decided randomly. This is shown by any decimal values in the players extra points.


