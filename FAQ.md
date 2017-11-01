Frequently Asked Questions

How come I am not seeing my function show up in the workbook or Function Wizard?
Most likely because you didn't load it. If you are running in the debugger, be sure to specify the full path to the Excel executable on your machine as the Command and "$(TargetPath)" as Command Arguments in Properties|Configuration Properties|Debugging.

When I open my .xll file in Excel nothing happens.
You need to set macro security to low in order for Excel to register your functions and macros. For Excel 2003 and earlier, go to Tools|Options... and click on Macro Security... in the Security tab. For Excel 2007 and later, from the Office Button click on Excel Options next to Exit Excel on the bottom right. On the Trust Center tab click on Trust Center Settings... On the Macro Settings tab choose Enable all macros.

How do I get my add-in to load every time I start Excel?
That is what the Excel Add-in Manager is for.

What is up with all those popups showing me errors and warnings?
The xll add-in library provides three levels of alerts: error, warning, and information. Hit Alt-F8 and type ALERT.FILTER to control which of these you see.
Image of ALERT.FILTER
Your choices are stored in the registry when you hit OK so they persist between Excel sessions. Hitting Cancel is equivalent to deselecting all alerts.

Why am I seeing a message box claiming Register failed for: X, where X is the Excel name of one of my functions?
You forgot to insert #pragma XLLEXPORT as the first line in the body of the corresponding C/C++ function.

What can I do if XLL.DOC is not generating my help file?
Look in the Debug directory. You should see a folder with the same name as the xll you are building. In that you will find .log files. Look through them in alphabetical order for error messages.

Why is malloc() failing while trying to allocate large amounts of memory?
You must compile with default char unsigned. Properties|C/C++|Command Line|Additional Options: /J

How come the last arguments in the constructor of AddIn end with an extra character?
It is a known feature that persists in the Excel 2010 SDK that the last character gets truncated.

I'm using Vista/Windows 7 and can't veiw the Macrofun.hlp file.
You need to install the Windows Help program. Just follow the directions

When I call my add-in function from Excel it returns #VALUE.
This can occur when there is a mismatch between the C signature you specify in the second argument to AddIn and the true signature.

Last edited Jul 4, 2011 at 10:32 AM by keithalewis, version 10