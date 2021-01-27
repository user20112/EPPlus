# EPPlus LGPL SkiaSharp Version
This was split from EPPlus Version 4.5 when it still used the LGPL liscence. warning Version5 uses the Polyform Noncommercial license
you can find this info at the link below [our website]( https://www.epplussoftware.com)
Create advanced Excel spreadsheets using .NET, without the need of interop.
EPPlus is a .NET library that reads and writes Excel files using the Office Open XML format (xlsx). 
EPPlus has the following dependencies
SkiaSharp (used as a replacement for system.drawing. this allows it to be used on ios android and uwp where that is not available).

## EPPlus supports:
* Cell Ranges 
* Cell styling (Border, Color, Fill, Font, Number, Alignments) 
* Data validation 
* Conditional formatting 
* Charts 
* Pictures 
* Shapes 
* Comments 
* Tables 
* Pivot tables 
* Protection 
* Encryption 
* VBA 
* Formula calculation 
* Many more... 
Note: this no longer supports underline and strike through as the old version did due to skiasharp limitations.
## Overview
This project started with the source from ExcelPackage. It was a great project to start from.
It had the basic functionality needed to read and write a spreadsheet.
Advantages over other:
EPPlus uses dictionaries to access cell data, making performance a lot better.
Complete integration with .NET 

## Support
This version should be identical to the old version with skiasharp replacements for the system.drawing components. therefore you can use the link below for most support but it may not be exactly the same.
All support is currently referred to [Stack overflow](https://stackoverflow.com/questions/tagged/epplus). 
A tutorial is available in the wiki and the sample project can be downloaded with each version. 
The old site at [Codeplex](http://epplus.codeplex.com) also contains material that can be helpful. 
Bugs and new feature requests can be added to the issues tracker. 

## License
The project is licensed under the GNU Library General Public License (LGPL). 
