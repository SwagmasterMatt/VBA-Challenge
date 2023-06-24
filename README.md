# VBA-Challenge

Getting Started:
Link to Excel File in One Drive: https://1drv.ms/x/s!Am4rPW8DSBZ4tm_ggvvR8tVY_-mv?e=gKEzKa
Ensure that GlobalsColor.bas and Main_Module2.bas are appropriately imported into the worksheet modules.
They should appear as "GlobalsColor" and "Main"
Navigate in the VBA editor in excel to the Tools - References and ensure the following are checked
- Visual Basic For Applications
- Microsoft Excel 16.0 Object Library
- OLE Automation
- Microsoft Office 16.0 Object Library
- Microsoft Scripting Runtime ** without this, the dictionaries will not function properly


GlobalsColor.Bas is a necessary module for the subroutine CallColors to function. It generates public variables and is a portable module I made for use in any VBA scripting.
Snips verify the start and end of the procedural generation of summaries. 
results for greatest values genereated by VBA and verified for accuracy with spot checks through manual calculations and excel Min / Max formuals in a separate workbook.

Main_Module2 uses a dictionary system and coverted date values so that the code is more versatile to allow for unordered data or data discrepencies.
Total runtime on a 16 core, 32 GB Ram system is around 93 seconds. Debug statements print times for each block.


Read Only Access
