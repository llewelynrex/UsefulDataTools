UsefulDataTools
===============

StringExtensions
----------------
The following extension methods are contained within this class:
* Left - A port of the VB.Net function to C# using Substring.
* Right - A port of the VB.Net function to C# using Substring.
* ToTrimmedString - Takes any object, runs ToString and at the same time trims the string.
* AddWhitespaceLeft - Creates a new string with the specified number of spaces to the left.
* AddWhitespaceRight - Creates a new string with the specified number of spaces to the right.

The following are some type conversions that converts to the specified type if possible:
* ToByte
* ToSByte
* ToInt
* ToUInt
* ToShort
* ToUShort
* ToLong
* ToULong
* ToFloat
* ToDouble
* ToDecimal
* ToBool
* ToChar
* ToDateTime

CsvOutputExtensions
-------------------
* ToCsv - Outputs the object as a CSV string converting all simple types to string.

ExcelOutputExtensions
---------------------
* ToExcel - Outputs an IEnumerable of type T, DataTable or DataSet as an Excel file, using the NetOffice library, by using the public properties and public fields  which have been found to be simple types and then writing them into an Excel sheet.
* ExcelOutputItem - An item that encapsulates the properties required to create an Excel spreadsheet.
* ExcelOutputCollection - A collection of ExcelOutputItems which, when exported, creates an Excel spreadsheet with a worksheet for each ExcelOutputItem.

Simple Types
------------
byte
sbyte
int
uint
short
ushort
long
ulong
float
double
decimal
bool
char
DateTime
byte?
sbyte?
int?
uint?
short?
ushort?
long?
ulong?
float?
double?
decimal?
bool?
char?
DateTime?
string
System.Enum
