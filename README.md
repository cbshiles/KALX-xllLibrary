This is a library for creating xll add-ins for Excel from 97 through 2014. It makes every feature of the latest Excel SDK available to you including the big grid, wide character strings, and asynchronous functions. It is the easiest way to integrate your C and C++, or even Fortran, code into Excel to achieve the highest possible performance. You can also generate native documentation using the same tool Microsoft uses for their help files.

If you need a small, fast, portable, and self contained way to extend Excel's functionality, this is the library for you. Just hand someone the xll and chm help file that you create and they are ready to go. No need to figure out what version of .Net they run, no Primary Interop Assemblies to worry about, no managed code that forces you to marshal data back and forth from Excel. There are also no automagic code generators, no proprietary markup languages to learn, and no wizards that hide things behind your back. Everything is just pure, modern, and readable C++.

NEW! Check out xllblog for the latest goodies in the pipeline.

INSTALLATION

Run setup.msi and project2010.msi . You will need to install htmlhelp.exe from Microsoft's HTML Help Workshop to build documentation.

OPERATION

Start Visual Studio 2010. This works with the free version of Visual C++ 2010 Express 
File > New > Project... select XLL AddIn Project. 
Right click on your project, Build. 
Right click on your project, Debug > Start new instance.

If all goes well, you should see XLL.FUNCTION show up under My Category in the Function Wizard. You can set breakpoints by clicking to the left of any line in the source code for xll_function in the file function.cpp.

To create HTML Help documentation, hit Alt-F8 then run XLL.DOC. Click on Help on this function in the Function Wizard to start the help viewer.

Function Wizard
INSTALLATION AND OPERATION VIDEO

CONGRATULATIONS!

You have just created and documented your first Excel add-in. See the Documentation to learn about all the cool stuff you can do with this library then check out the Related Projects to see examples of how to use this library.

CREDITS

As with any non-trivial endeavor, this library leveraged off of the knowledge of others. I give special thanks to the people who have spent many hours over the years spelunking the deepest caves of the Excel SDK: Govert van Drimmelen, Mark Joshi, Jérôme LeCompte, Laurent Longre. Kudos to Eric Woodruff for his excellent Sandcastle Help File Builder that made it possible to carve out the subset of Sandcastle used here. My sincere apologies to anyone I've neglected to mention.

TESTIMONIALS

'Previously, we used VBA but it was so slow - by more than 70%!' - Ki-Hwan Bae

'Thanks again for the awesome library! I've been testing out a couple different methods/libraries for creating an Excel add-in, and so far using the XLL approach paired with your library is my favorite because it gives you all the low level access without hiding everything behind magic button code generators (like all the commercial products I've tested out seem to) while keeping the C++ complexity in check with the nice wrapper utility classes.' - Colin Rodriguez

'Thank you. I need high performance computing as well. I saw a lot of in-memory copying in XLW which makes me wary.' - Candy Chiu

 

Last edited Feb 9, 2015 at 11:43 AM by keithalewis, version 96

