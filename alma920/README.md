# alma920 PowerShell script

**Name:** alma920.ps1 <br/>
**Created:** 2018-12-04 by V.M. Downey for Houghton Technical Services <br/>
**Description:** Windows PowerShell v. 1 script to create a text string for the Alma 920 field and copy the string to the clipboard <br/>

## How to use
1. Create a new field in the Alma holdings record (F8)
1. Type "920" for the field and "11" for the indicator values
1. Double click on the alma920 shortcut (or use shortcut keys, if assigned)
1. Select a cataloging type ($$a) value and a curatorial department ($$x)
1. Click OK to save the 920 string to the clipboard
1. Paste string into 920 field

## Setup (Windows 10)
1. Save alma920.ps1 to workstation
1. Create a shortcut to alma920.ps1 (right click on file and select Create Shortcut)
1. Move shortcut to Desktop (and, optionally, rename shortcut)
1. Right click on shortcut and select Properties
1. Select Shortcut tab
1. Copy and paste following text into Target box: powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File alma920.ps1
1. (Optional) Assign shortcut keys using the Shortcut key box
1. Change Run value to "Minimized"
1. Click Apply
1. Click OK

Note: The blank $$a can be left in the 920 field. When the record is saved, the blank subfield will be deleted.

## Notes
* Default cataloging type ($$a) value is "a" (accession) and default curatorial department ($$x) value is "modbm" (Modern Books & Manuscripts).
* The date ($$d) and cataloger ($$f) values are computer-generated. The date is today's date; the cataloger is from the computer username.
