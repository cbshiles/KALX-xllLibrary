This is a library for creating xll add-in's for versions of Excel before
and after Excel 2007.  It makes every feature of the Excel 2010 SDK
available to you, including 64-bit builds. (Not yet!)

INSTALLATION

Run setup.msi and project.msi from http://xll.codeplex.com.

OPERATION

In Visual Studio 2010 create a new project of type XLL AddIn Project.
Right click on project, Build.
Right click on project, Debug > Start new instance.

Set breakpoints by clicking to the left of the line.

USAGE

The naming convention is the same as in the Microsoft SDK: just add a
'12' at the end of names.

If you add an 'X' instead of '12' you can use the same source
code to build add-ins for either version.  Just define EXCEL12
before including "xll/xll.h" to create add-ins for the latest
versions of Excel.  We use the same UNICODE conventions as
Microsoft. http://msdn.microsoft.com/en-us/goglobal/bb688113.aspx

TROUBLESHOOTING
http://support.microsoft.com/default.aspx?scid=kb;en-us;Q280504
http://support.microsoft.com/kb/948615