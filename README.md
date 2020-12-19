# wellsrPRO
| Note    	| Description                                                                                                                                             	|
|---------	|---------------------------------------------------------------------------------------------------------------------------------------------------------	|
| #1 	| Clicking wellsrPRO.xlam to load the add-in may cause macros to become disabled! Follow the installation instructions below for proper installation! 	|

*********************************************************************************************
### To prevent the add-in from disappearing when you restart Excel, follow these steps. 

1) Close Excel
2) Unzip the wellsrPRO zip file to extract the contents
3) Right-click the wellsrPRO.xlam file in Windows Explorer
4) Click Properties
5) On the General tab, you may see a message that states: "This file came from another computer
   and might be blocked to help protect this computer."
   a) If you see this message, click Unblock 
   b) Click OK
   c) Proceed to installation instructions

If this still doesn't work, try navigating to %APPDATA%\Microsoft\Excel\XLSTART and placing 
the wellsrPRO.xlam file here.

If this still doesn't work, Launch Excel, Open your VBA Editor, create a New Module, then
run the following macro:
    Sub OpenXLSTART()
        Shell "explorer.exe " & Application.StartupPath, vbNormalFocus
    End Sub

Place your wellsrPRO file in the folder that opens (debug.print application.StartupPath)


*********************************************************************************************
### Installing wellsrPRO on Excel 2010 and newer

Installing wellsrPRO on Excel 2010 and newer:
1) If you haven't done so already, unzip the wellsrPRO zip file to extract the contents
2) Open Excel
3) Click the File tab
4) Click Options
5) Click the Add-Ins category on the left bar of the Excel Options Dialog Box
6) In the Manage box near the bottom, click Excel Add-ins, and then click Go. The Add-Ins dialog box appears.
7) Click Browse
8) Navigate to the directory where you extracted wellsrPRO.zip
9) Select wellsrPRO.xlam, and then click OK. You may be asked to copy the add-in to your local Microsoft AddIns directory. This is optional.
10) Make sure the checkbox next to "wellsrPRO" is checked, then click OK.
11) If the add-in keeps disappearing when you restart Excel, follow the directions at the top of this file

*************************************************************************************************
<i>You cannot run update.vbs manually. Attempting to do so will generate an error. This script is an internal script used by wellsrPRO to check for updates. You can ignore it. :-) </i>
*************************************************************************************************

Visit wellsr.com for more information and navigate the "Help and Info" section on the top ribbon for assistance with the add-in. 

Meet the developer: 
     [@ryanwellsr](https://twitter.com/ryanwellsr)

