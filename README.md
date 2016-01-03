# Exporter
Export to Excel,Word , Pdf, Html,CSV

[![Build status](https://ci.appveyor.com/api/projects/status/w4w6k0kxu2cide0m?svg=true)](https://ci.appveyor.com/project/ignatandrei/exporter)




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

//or you can use the factory
            data = ExportFactory.ExportData(new List<Person>() {p}, ExportToFormat.Excel2007);
            File.WriteAllBytes("b.xlsx", data);
            Process.Start("b.xlsx");
