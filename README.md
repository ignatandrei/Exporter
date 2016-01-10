# Exporter

[![Join the chat at https://gitter.im/ignatandrei/Exporter](https://badges.gitter.im/ignatandrei/Exporter.svg)](https://gitter.im/ignatandrei/Exporter?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge&utm_content=badge)
Export to Excel,Word , Pdf, Html,CSV

The data could be a C# class or a JSON string or a CSV string or a DataTable

NuGet package at https://www.nuget.org/packages/ExporterWordExcelPDF/ 

<a href="https://www.nuget.org/packages/ExporterWordExcelPDF/"><img src="https://img.shields.io/nuget/v/ExporterWordExcelPDF.svg"></img></a>


The source code has tests and a console project for easy testing the package


[![Build status](https://ci.appveyor.com/api/projects/status/w4w6k0kxu2cide0m/branch/master?svg=true)](https://ci.appveyor.com/project/ignatandrei/exporter/branch/master)

Code examples in C# / JavaScript at <https://github.com/ignatandrei/Exporter/wiki>

Demo online at <http://exporter.azurewebsites.net/>

You can contribute to the project - read <https://github.com/ignatandrei/Exporter/wiki/Help-the-project>


PS: Just to make you an idea, this can be the code to export to Excel
```csharp
    List<Person> listWithPerson  = ... //obtained from database
    var export=new ExportExcel2007<Person>();
    var data = export.ExportResult(listWithPerson);
    File.WriteAllBytes("a.xlsx", data);
    Process.Start("a.xlsx");
```
(Do not forget 
```csharp    
using ExporterObjects;
using ExportImplementation;  
```
)


But is better to read the [Wiki]<https://github.com/ignatandrei/Exporter/wiki>
