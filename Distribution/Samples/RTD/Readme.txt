To make the RTD examples work, make a copy of ExcelDna.xll in this directory, rename it to TestRTD.xll, and open in Excel.

Some functions from the TestRTD.dna to try:
(By default Excel will update every 2 seconds.)

A ticking clock
---------------
=WhatTimeIsIt()

Live data from a file 
------------------------
(change the contents of the Test.xml file to see Excel update).
=GetEurOnd()
=GetTestItem("EUR/DEPO/ON")
