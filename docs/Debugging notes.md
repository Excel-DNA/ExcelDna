---
layout: page
title: Debugging Notes
---

## Debugging user code in your add-in

I don't usually have issues debugging the managed add-in's user-defined functions and commands.

There is a trick needed to set  up the debugging when using the Express editions - the project file must be edited by hand to set the native executable to run. This is discussed here: [https://groups.google.com/group/exceldna/browse_thread/thread/6767ac80f2bb3f11](https://groups.google.com/group/exceldna/browse_thread/thread/6767ac80f2bb3f11).


## Debugging the Excel-DNA integration libraries


_This is from a discussion on the Google group - [https://groups.google.com/group/exceldna/browse_frm/thread/b2d93b2e5e986bbd](https://groups.google.com/group/exceldna/browse_frm/thread/b2d93b2e5e986bbd)_

Debugging the managed Excel-DNA integration libraries can be a bit tricky sometimes. 

When recompiling (only for the Debug - Win32 configuration), you'll see the ExcelDna.Loader.dll and ExcelDna.Integration.dll copied to the output (Source\ExcelDna\Debug) directory. Then, when the add-in is loaded from there, these assemblies are loaded from the files and not the packed versions. So you need not fiddle with the packing for the debugging to work - you just need those assemblies as files next to the .xll - then the files are loaded and the packed versions ignored. They have to be up-to-date  though, else the debugger won't match your code with the loaded .dlls.

The following works for me:

* Make sure you rebuild everything - when you set up 'ExcelDna.Integration' for debugging, the ExcelDna project might not rebuild, so the ExcelDna.Integration in the (Source\ExcelDna\Debug) directory might not be updated. When in doubt I just press Build->Batch Build->Rebuild All, and then I'm sure it's OK.

* I start Excel on its own, then attach the debugger, then load the add-in. This seems more reliable than starting the debugger 

* Another trick I have use sometimes is to put a Debug.Assert(false) somewhere in the loading path - this allows you to select the debugger when the assertion failure breaks the process.

