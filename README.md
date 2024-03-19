In Excel have you ever wanted to have database like functionality be able to run from within VBA. You could link to an MDB which can be slow, and a bit cumbersome. You could also custom write functions to work in the Excel worksheet. If you really want speed you could use an Array in memory, but it can be a bit tiresome writing the code to manipulate the array. ExcelDataSet takes the idea of the C# DataSet and implements an in memory data manipulation framework. It is fast, based on the Windows Scripting Dictionary object. It also outputs to XML.

Crucially it is very easy to load Excel data into the ExcelDataSet tables via an array, manipulate the data in memory and then output it back into an array and the spreadsheet if needed.

To most quickly get an idea of what this is all about and how it can be used, you can download the cDataSet.xlsm file and run the Demo. Have a look at the code in the Demo module to see how it calls the Dataset classes
