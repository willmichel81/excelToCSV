# excelToCSV
The <code>xlsxToCsv</code> package contains the ability to convert .xlsx file to a .csv file that is more consumable by many other packages/programs/scripts.<br /><br />

The source code for this package is available at https://github.com/willmichel81/excelToCSV.<br /><br />

The importable name of the package is excelToCSV<br /><br />

<h2>Import into your code</h2>
<code>
>>> from excelToCSV import excelToCSV
</code>
<br /><br />

<h2>Test to make sure package is connecting</h2>
<code>
>>> pip install xlsxToCsv <br />
>>> python <br />
>>> from excelToCSV import excelToCSV <br />
>>> test = excelToCSV(C:/complete/path/to/file/example.txt) <br />
>>> print(test)
</code><br />
 
 If you have succesfully installed xlsxToCsv then you should get the following results from the above code <br />
 <code>
 >>> "('C:/complete/path/to/file/', '.txt')"
 </code>

<h2>Short Summery</h2>
Takes user specified range data (ex. A1:A16) data from sheet(s)(ex. "MAIN") in side of .xlsx and then converts data to json dict and/or csv.
