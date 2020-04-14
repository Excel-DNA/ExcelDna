---
layout: page
title: What and why - An introduction to .NET and Excel-DNA
---

Microsoft publishes a Software Development Kit (SDK) for Excel, that describes how to make an 'Excel add-in'. These are similar to .xla add-ins, but the code would typically be written in C or C++, and is compiled to binary file with a .xll extension. Such an Excel add-in is typically called an ".xll". Internally, it is just a normal Windows .dll file with a few pre-arranged exports so that Excel and the add-in can hook up.

Xll add-in have some advantages over .xla add-ins developed in VBA. They can define user-defined worksheet functions (UDFs) which run very fast and which can integrate with newer features like multi-threaded calculation in Excel 2007+, and asynchronous calculation in Excel 2010+. A disadvantage of xll add-ins is that they are hard to develop. Typically C or C++ is used, and there are some toolkits and a nice book to help you, but it's still hard.

The .NET Framework (normally just called .NET) is a (twelve-year old) software framework from Microsoft, including the (Java-like) C# language and an updated version of Visual Basic, extensive standard libraries, a runtime environment (libraries that manage execution of your code) and an intermediate 'assembly' language with corresponding just-in-time compiler. The .NET languages and runtime environment is often called 'managed'. So a .NET library would be called a 'managed' library, as opposed to a library compiled from C / C++, which would be a 'native' library. Wikipedia can tell you a lot more about .NET. It has become the standard development environment for corporate software development on the Windows platform.

The Microsoft development tool (giving you the compilers and Integrated Development Environment (IDE)) associated with .NET is called Visual Studio. What you might think of as the 'real' Visual Basic (the last version was VB6) was upgraded to become a language as part of .NET, often called VB.NET, and Visual Studio is the standard IDE for developing VB.NET applications. There are free editions of Visual Studio, and then a range of paid for editions with more and more features.

There is some support in Visual Studio for making Office add-ins using .NET, with a library called Visual Studio Tools for Office (VSTO). However, initially (ten years ago) VSTO had many complications with deployment, and particular for Excel has serious limitations - that UDFs could not be created.

So for Excel there was a problem - how to use .NET to create full-featured and high-performance add-ins (meaning .xll add-ins). There was a commercial solution called ManagedXLL when I looked around in 2004, but it was too expensive to be useful to me. So I started an open-source project called Excel-DNA (the 'DNA' stands for DotNet for Applications, as opposed to Visual Basic for Applications).

The main Excel-DNA sites are [http://excel-dna.net](http://excel-dna.net) and [https://excel-dna.github.io](https://excel-dna.github.io).
Now (after ten years) Excel-DNA is mature and widely used as the standard .NET to Excel integration tool. But in practice, it is most useful for developers already using .NET and Visual Studio who want to make high-performance Excel add-ins. 

To be sure, there are other ways of making .xll add-ins, including a tool called PyXll that is similar to Excel-DNA, but for the Python language rather than the .NET Framework.

I hope in the next few years to make it easier for Excel VBA users to upgrade to using VB.NET and Excel-DNA. It has proven harder than I expected, as there is a steep learning curve, and the advantages are uneven and very dependent on the programmer's background.

One guide, written from a VBA user's perspective, that you might look at to see whether you might be interested is a porting guide written by Patrick O'Beirne: [http://sysmod.wordpress.com/2012/11/06/migrating-an-excel-vba-add-in-to-a-vb-net-xll-with-excel-dna-update/](http://sysmod.wordpress.com/2012/11/06/migrating-an-excel-vba-add-in-to-a-vb-net-xll-with-excel-dna-update/)

So if you have existing VBA code (mainly worksheet functions) that you want to run fast in newer versions of Excel on modern hardware, Excel-DNA is a good tool. You'd port your VBA code to VB.NET (using the free version of Visual Studio) and use Excel-DNA to glue that back to Excel. I'm more than happy to help users learning how to do this on the Google group (or sometimes directly), but there is a steep and long learning curve once you put your foot outside the familiarity of VBA.
I believe it's worth the effort.
