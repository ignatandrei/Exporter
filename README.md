# Exporter
Export to Excel,Word , Pdf, Html,CSV

The data could be a C# class or a JSON string

NuGet package at https://www.nuget.org/packages/ExporterWordExcelPDF/ 

<a href="https://www.nuget.org/packages/ExporterWordExcelPDF/"><img src="https://img.shields.io/nuget/v/ExporterWordExcelPDF.svg"></img></a>


Demo at http://exporter.azurewebsites.net/


[![Build status](https://ci.appveyor.com/api/projects/status/w4w6k0kxu2cide0m/branch/master?svg=true)](https://ci.appveyor.com/project/ignatandrei/exporter/branch/master)


Usage:

var p = new Person { Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/" , CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };

            Export<Person> export = new ExportHtml<Person>();
            var data = export.ExportResult(new List<Person>() { p });            
            File.WriteAllBytes("a.html",data);
            Process.Start("a.html");

            export=new ExportExcel2007<Person>();
            data = export.ExportResult(new List<Person>() { p });
            File.WriteAllBytes("a.xlsx", data);
            Process.Start("a.xlsx");


            export = new ExportWord2007<Person>();
            data = export.ExportResult(new List<Person>() { p });
            File.WriteAllBytes("a.docx", data);
            Process.Start("a.docx");


            export = new ExportPdfiTextSharp4<Person>();
            data = export.ExportResult(new List<Person>() { p });
            File.WriteAllBytes("a.pdf", data);
            Process.Start("a.pdf");
//using factory

            data = ExportFactory.ExportData(new List<Person>() {p}, ExportToFormat.Excel2007);
            File.WriteAllBytes("b.xlsx", data);
            Process.Start("b.xlsx");
            
            
            
             //export from JSON
            string dataJson = @"[
    { 'Name':'Andrei Ignat', 
        'WebSite':'http://msprogrammer.serviciipeweb.ro/',
        'CV':'http://serviciipeweb.ro/iafblog/content/binary/cv.doc'        
    },
{ 'Name':'Andrei Ignat', 
        'WebSite':'http://msprogrammer.serviciipeweb.ro/',
        'CV':'http://serviciipeweb.ro/iafblog/content/binary/cv.doc'        
    }
]";

            data= ExportFactory.ExportDataJson(dataJson, ExportToFormat.Excel2007);
            File.WriteAllBytes("b1.xlsx", data);
            Process.Start("b1.xlsx");
            
More details at <https://github.com/ignatandrei/Exporter/wiki>


You can also buy me a beer by donating at 

[PayPal](https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=ignatandrei@yahoo.com&item_name=Exporter&item_number=GitHub)
