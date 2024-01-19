# EUVATChecker
A program that checks if VAT numbers in an Excel spreadsheet are valid and active using a command line program. 

## Prerequisits 
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
Place the xlsx in the same folder as the .exe of the EUVATChecker and then execute the .exe.

The code is in C# as the library used was using the .NET framework.
