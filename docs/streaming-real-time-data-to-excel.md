---
layout: post
title: "Streaming real-time data to Excel"
date: 2013-10-07 23:13:00 -0000
permalink: /2013/10/07/streaming-real-time-data-to-excel/
categories: uncategorized, examples, excel, rtd, rx
---
Gert-Jan van der Kamp has posted a very nice end-to-end example on [CodeProject][codeproject-streaming], showing how to create a WCF service and Excel-DNA add-in to stream real-time data into Excel.

The example uses to use the Reactive Extensions support in Excel-DNA v. 0.30 to push the data to an Excel UDF (using Excel's RTD mechanism behind the scenes), together with a Duplex WCF service providing the data.

There was also this [CodePlex discussion][codeplex-460904] about the Excel ThrottleInterval option, which trades off the real-time update frequency against stability of the Excel calculation.

[codeproject-streaming]: http://www.codeproject.com/Articles/662009/Streaming-realtime-data-to-Excel
[codeplex-460904]: https://exceldna.codeplex.com/discussions/460904
