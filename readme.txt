Bulk Copy Process by Silverlance Productions

BCP is a method to transfer large amounts of data from a flat file into SQL Server VERY quickly.  BCP is a functionality of the low-level DBLib environment.  Unfortunately the modern data components (ADO, DAO, RDO) don't support bulk copy.  

The only way to get to BCP in the past was to either use the BCP.exe tool shipped with SQL Server, or to write some serious low level BCP type stuff.

Then one day the good folks of Microsoft came up with VBSQL.ocx.  It is a wrapper control for DBLib.  Now this wrapper control is a leap forward for us VB Developers, but the documentation for it .. well... sucks.

I wrote slpBCP.dll as a wrapper for VBSQL.ocx and DBLib.  It will allow you the developer to reference the DLL in your project, Create a reference to the object, set a few properties, call a routine... and bulk copy data into your SQL server.

Things You Need To Know:

The current methodology of my shop is that all flat files are copied into matching DB tables where all the field types are char().  There is no exception to this rule, so that's all that slpBCP supports.  Our process is to load the data into the SQL Server, then use stored procedures and the servers horsepower to manipulate the data once it's in the DB.

Everything that we do with flat files is fixed width, so that's all slpBCP supports.  No comma delimited, no tab delimited.. Nothing but flat fixed width ASCII files.

slpBCP only supports Bulk copy into the SQL Server, not out of.  Extracting data from SQL Server is planed for future enhancements.. but for now it's just on the wish list.

Since this DLL wraps up VBSQL.ocx (which is part of the SQLServer PTK), you need to have it on your system, and you need to distribute it. With your application.  I have no included it in this source code distribution since I wanted to keep the zip file small.

Last but not least, you will need the DBLib files.  There is more information of which files you will need in the PTK, so I won't go into it here..

Properties:

RowsCopied = Long - Read only, returns the number of records loaded into the 
SQL Server

TruncateData = Boolean - Sets/Returns whether or not slpBCP will truncate the 
target table before doing it's Bulk Copy Process

TargetTable = String - Sets/Returns the name of the table to copy into

TargetDB = String - Sets/Returns the name of the database to copy into

TargetServer = String - Sets/Returns the name of the server to copy into

Login = String - Sets/Returns the login ID 

Password = String - Sets/Returns the Login Password

SourceFile = String - Sets/Returns the Source File name and path to be copied 
into the SQL Server

BCPBatchCount = Integer - Sets/Returns the number of rows to copy into the 
server at a time.  The default value for this property is 1000 and does not need to be changed.

BCPColumnCount = Integer - Sets/Returns the number of columns in the source 
file and in the target table.  This property will eventually go away, but I was in a bit of a hurry.. so there it is :)


Methods:

StartBulkCopy (Returns boolean) - Always returns true.. if there is ever an error in the slpBCP process, it will raise a trappable error.

Errors:

ieLogin = 5001 = "DBLib could not allocate login record"
ieBCPLogin = 5002 = "DBLib count not Enable BCP for this login"
ieSQLOpen = 5003 = "DBLib count not open connection to server"
ieSourceFileNotFound = 5004 = "SourceFile Not found on system"
ieBCPInitFailed = 5005 = "DBLib unable to Init the BCP functionality"
ieBCPControl = 5006 = "DBLib unable to set Batch Row Count"
ieBCPColumns = 5007 = "DBLib unable to set the BCP Column Count"
ieBCPColFormat = 5008 = "DBLib unable to set the Column Format"
ieBCPExec = 5009 = "DBLib Execute BCP Failed"
ieCMDFail = 5010 = "DBLib Failed to put the SQL Command on the Stack"
ieSQLExec = 5011 = "DBLib SQL Execute Failed"
ieSQLResults = 5012 = "DBLib Get of SQL Results failed"


Example of use:

' Init our BCP Class
Set cBCPLoad = New slpBCP.clsBCPLoad
cBCPLoad.BCPColumnCount = 5
cBCPLoad.SourceFile = "c:\Temp.txt"
cBCPLoad.TargetTable = "tblLoadtemp"
cBCPLoad.Login = "sa"
cBCPLoad.Password = "hidden"
cBCPLoad.TargetDB = "Pubs
cBCPLoad.TargetServer = "sqlserver1
cBCPLoad.StartBulkCopy
Set cBCPLoad = Nothing


Disclaimer:
This code is provided as is.  Your use of this code either in whole or in part releases the author of this code from any and all liability from any damage including but not limited to loss of data, loss of life, damage to hardware, software, or psyche.

Notes:
I wrote this to help me in my every day life.. if you like it, let me know.. if you don't like it.. let me know.. if you find a bug, let me know.. if you think it's flawless, let me know.

Jason K. Monroe - datacop@iei.net
