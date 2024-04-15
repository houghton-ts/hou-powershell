# Alma 920 field PowerShell script

**Name:** alma920.ps1 <br/>
**Created:** 2018-12-05 by V.M. Downey for Houghton Technical Services <br/>
**Description:** Windows PowerShell v. 1 script to create a text string for the Alma 920 field and copy the string to the clipboard <br/>

This Windows PowerShell script creates a text string for the Alma 920 field and copies the string to the clipboard for a user to paste in the Alma Metadata Editor. Harvard University libraries use this local MARC holdings field for recording cataloging statistics.

## How to use
1. Create a new field in the Alma holdings record (F8)
1. Type "920" for the field and "11" for the indicator values
1. Double click on the alma920 shortcut (or use shortcut keys, if assigned)
1. Select cataloging type ($$a), project name ($$p), and curatorial department ($$x). Leave $$p blank if not applicable.
1. Click OK to save the 920 string to the clipboard
1. Paste string into 920 field in Alma

## Installation (Windows 11)
1. Remove previously downloaded alma920.ps1 file and shortcuts (skip to step 2 if installing for the first time)
   * Locate the previously downloaded .ps1 file
     * If you have a shortcut, right click on shortcut and select "Properties"
     * Select Shortcut tab
     * Copy "Start in" folder path
     * Navigate to the folder in Windows Explorer
   * Delete the file
   * Delete the Desktop shortcut
   * Save new or updated alma920.ps1 to your computer by right-clicking on the file above and choosing "Save link as"
1. Create a shortcut to alma920.ps1
    * Right click on file
    * Select "Show more options"
    * Select "Create shortcut"
    * Move shortcut to Desktop (and, optionally, rename shortcut)
    * Right click on shortcut and select "Properties"
    * Select Shortcut tab
    * Copy and paste following text into Target box: powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File alma920.ps1
      * -ExecutionPolicy Bypass allows the script to run without being blocked or triggering security warnings or prompts
      * -NoProfile does not load the PowerShell profile
      * -WindowStyle Hidden hides the command prompt window
1. (Optional) Assign shortcut keys using the Shortcut key box
1. Change Run value to "Minimized"
1. Click Apply
1. Click OK
1. Double click on the shortcut to run the script

Note: Blank subfields will be deleted when the Alma record is saved.

## Notes
* [Houghton Implementation of the 920 Field](https://wiki.harvard.edu/confluence/x/9rDkDQ)
* [920 PowerShell Script and Documentation](https://wiki.harvard.edu/confluence/x/S7HkDQ)
* The date ($$d) and cataloger ($$f) values are script-generated. The date is today's date; the cataloger is from the computer username.

## Development History
Created: 2018-12-05 by V.M. Downey for Houghton Library Technical Services <br/>
Updated: 2022-06-14 by V.M. Downey to add new $$x codes and enlarge box size <br/>
Updated: 2024-03-12 by V.M. Downey to add $$a r for reparative description and $$p for projects <br/>
Updated: 2024-03-18 by V.M. Downey to change numbered label/list box variables to descriptive variable names <br/>
