ComServer sample
----------------
Shows how Excel-DNA can serve .NET classes in an add-in library from within the same AppDomain as the add-in/ This allows easy interaction between UDFs and VBA code, and give another way to integrate .NET code into Excel.

To run the sample:
1. Run BuildSample.bat to compile the libraries, create and registier the ComServerPacked.xll add-in.
2. Open the ComServerTest.xls and enable macros.
3. The function on the worksheet won't work unless the .xll is loaded.
4. Go to the VBA editor (Alt+F11) and run the two test Subs.
5. Load the .xll into Excel to register the functions.
6. Check that the worksheet function calculates and uses the value set in the VBA code.
