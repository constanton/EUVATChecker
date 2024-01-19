# EUVATChecker
A program that checks if VAT numbers in an Excel spreadsheet are valid and active using a command line program. 

## Prerequisites 
To build this code you need the [vies-dotnet](https://github.com/zapadi/vies-dotnet/tree/master) library and the [OpenXML-SDK](https://github.com/dotnet/Open-XML-SDK) library.
Both are available in Visual Studio via NuGet Package manager (use vies-dotnet-api and DocumentFormat.OpenXml).

## How to use (once the code is compiled)
### Step 1
Create an xlsx file with an Excel spreadsheet making sure the VAT numbers are in Column B and the name of the VAT number holder is in Column C. The first row is ignored so you can keep it as a header for your table. For example:
| ID      | VAT Number | VAT Number Holder     | Description |
| ----------- | ----------- |----------- | ----------- |
| 1   | FR12345678901       |Desireless & Co. SARL  | Voyage Voyage       |
| 2   | DE123456789         |Kraftwerk & Co. GmbH | Radioactivity        |
| 3   | DK99999999          |Raveonettes & Co. ApS | She owns the streets |

### Step 2
Save the Excel file as vat.xlsx

### Step 3
Place the vat.xlsx file in the same folder as the .exe of the EUVATChecker and then execute the .exe.

### Good to know
* The program might fail if the EU VIES server is busy.
* If it's not busy then you might be able to even use 2000 rows in the spreadsheet.
* The code is in C# as the library used was using the .NET framework. If anyone has found a more recent Python library, please let me know.

### Licence
The code is under the MIT License.
