using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExporterObjects;
using ExportImplementation;

namespace ExporterConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var p = new Person { Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/" , CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };

            Export<Person> export = new ExportHtml<Person>();
            var data = export.ExportResult(new List<Person>() { p });            
            File.WriteAllBytes("a.html",data);
            Process.Start("a.html");

            export = new ExportExcel2007<Person>();
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
            data = ExportFactory.ExportData(new List<Person>() { p }, ExportToFormat.Excel2007);
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

            data = ExportFactory.ExportDataJson(dataJson, ExportToFormat.Excel2007);
            File.WriteAllBytes("bJson.xlsx", data);
            Process.Start("bJson.xlsx");


            //or from CSV
            var dataCSV = new List<string>();
            dataCSV.Add("Name,WebSite,CV");
            dataCSV.Add("Andrei Ignat,http://msprogrammer.serviciipeweb.ro/,http://serviciipeweb.ro/iafblog/content/binary/cv.doc");
            dataCSV.Add("Andrei Ignat,http://msprogrammer.serviciipeweb.ro/,http://serviciipeweb.ro/iafblog/content/binary/cv.doc");

            data = ExportFactory.ExportDataCsv(dataCSV.ToArray(), ExportToFormat.Excel2007);
            File.WriteAllBytes("bCSV.xlsx", data);
            Process.Start("bCSV.xlsx");

            var dta = new DataTable("andrei");
            dta.Columns.Add(new DataColumn("ID", typeof(int)));
            dta.Columns.Add(new DataColumn("Data", typeof(string)));
            
            dta.Rows.Add(1, "test 1 ");
            dta.Rows.Add(2, "test 2 ");
            dta.Rows.Add(3, "test 3 ");

            data = ExportFactory.ExportDataFromDataTable(dta, ExportToFormat.Excel2007);
            File.WriteAllBytes("dta.xlsx",data);
            Process.Start("a.xlsx");

            //advanced - modifying templates - in this case , added Number
            export = new ExportExcel2007<Person>();
            //File.WriteAllText("Excel2007Header.txt",export.ExportHeader);
            //File.WriteAllText("Excel2007Item.txt", export.ExportItem);
            //File.WriteAllText("Excel2007Collection.txt", export.ExportCollection);
            export.ExportHeader = File.ReadAllText("Excel2007Header.txt");
            export.ExportItem = File.ReadAllText("Excel2007Item.txt");
            export.ExportCollection = File.ReadAllText("Excel2007Collection.txt");
            data = export.ExportResult(new List<Person>() { p });
            File.WriteAllBytes("WithId.xlsx", data);
            Process.Start("WithId.xlsx");

        }
    }
}
