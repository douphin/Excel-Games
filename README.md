# Excel-Games
A Personal Project of trying to code simple games into excel using vba. Currently trying to get my first game, chess, running entirely in VBA. I'm still a little ways away from it being playable but I'm confident Its doable.

## How to Play
( Only Chess right now, and it's not fully working, but if you want to check out my progress, follow these steps )

### Method 1: Download PreMade Macro-Enabled Worksheet
 1. Choose the game you want to play
 2. Open the corresponding folder
 3. Click on the file with the .xlsm extension
 4. The workbook can now be downloaded using the `Downlaod Raw File` button
 5. The workbook now needs to be unlocked, follow instructions found [here](https://support.microsoft.com/en-us/topic/a-potentially-dangerous-macro-has-been-blocked-0952faa0-37e7-4316-b61d-5b5ed6024216) to do so
 6. Open the workbook and enjoy

### Method 2: Copy and Paste Code
1. Choose the game you want to play 
2. Open the corresponding folder
3. Click the file with the .vb extension
4. Copy all of code in the file, this can be done with the copy button to the top right of the code, or by clicking into the code, pressing `ctrl+a` then `ctrl+c` .
5. Now we'll paste the code into Excel
6. Open a desktop instance of Excel
7. Select the `View` tab found along the top menu bar
8. Select the `Macros` button found on the far right inside the `View` tab. (Click `View Macros`)
9. Create a new Macro by typing a name for it and clicking `Create` (the name doesn't matter, it'll be overwritten by time we're done)
10. At this point a new screen should have popped up showing the macro coding environment
11. In the coding environment, delete the words *Sub 'MacroName'* and *End Sub*
12. Now paste in all of the code copied earlier.
13. The macro coding environment can now be exited by pressing the close button on the window labled *Microsoft Visual Basic for Applications*
14. We're almost there, go back and follow steps 7 and 8 to re-open the macro menu
15. There should now be a whole host of macros listed, find the one labled `NewBoard`, select it and click `Run`
16. Everything to play should now be set up, Enjoy.

### Additional Info
If you used method 2 to copy and paste the code, and you want to save the workbook, you need to make sure you save it as a Macro-Enabled Workbook (.xlsm) not the standard (.xlsx) otherwise it won't save any of the macros.

If you are interested in writing your own macros, I would recommmend activating developer setting inside of excel, to do this:
1. Click `File` in the top left of Excel
2. Then go down and click `Options`
3. Then in the menu that pops up find the tab that says `Customize Ribbon` and click that
4. Now, on the right, find and check the box labeled `Developer`

