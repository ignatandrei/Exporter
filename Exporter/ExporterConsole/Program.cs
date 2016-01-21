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
            Console.WriteLine("Please choose:");
            Console.WriteLine("1. See all exports");
            Console.WriteLine("2. See advanced - modify template to add number");
            var option = Console.ReadLine();
            switch (option)
            {
                case "1":
                    SeeAllExport();
                    return;
                case "2":
                    Advanced_ModifyTemplates();
                    return;
                    
            }
            Console.WriteLine("demo finished");
            
           

        }

        static bool writeAndStartFile(string fileName, byte[] dataBytes)
        {
            if (File.Exists(fileName))
            {
                try
                {
                    File.Delete(fileName);
                }
                catch (Exception)
                {
                    return false;
                }
            }
            File.WriteAllBytes(fileName, dataBytes);
            Process.Start(fileName);
            return true;

        }
        static void SeeAllExport()
        {
            var p = new Person { Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/", CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };
            var p1 = new Person { Name = "you", WebSite = "http://yourwebsite.com/" };
            var list = new List<Person>() { p, p1 };
            Export<Person> export = new ExportHtml<Person>();
            var data = export.ExportResult(list);
            if(!writeAndStartFile("a.html", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete a.html");

            export = new ExportExcel2007<Person>();
            data = export.ExportResult(list);
            if(!writeAndStartFile("a.xlsx", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete a.xlsx");


            export = new ExportWord2007<Person>();
            data = export.ExportResult(list);
            if(!writeAndStartFile("a.docx", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete a.docx");


            export = new ExportPdfiTextSharp4<Person>();
            data = export.ExportResult(list);
            if(!writeAndStartFile("a.pdf", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete a.pdf");


            //or you can use the factory
            data = ExportFactory.ExportData(list, ExportToFormat.Excel2007);
            if(!writeAndStartFile("b.xlsx", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete b.xlsx");

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
            if(!writeAndStartFile("bJson.xlsx", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete bJson.xlsx");


            //or from CSV
            var dataCSV = new List<string>();
            dataCSV.Add("Name,WebSite,CV");
            dataCSV.Add("Andrei Ignat,http://msprogrammer.serviciipeweb.ro/,http://serviciipeweb.ro/iafblog/content/binary/cv.doc");
            dataCSV.Add("Andrei Ignat,http://msprogrammer.serviciipeweb.ro/,http://serviciipeweb.ro/iafblog/content/binary/cv.doc");

            data = ExportFactory.ExportDataCsv(dataCSV.ToArray(), ExportToFormat.Excel2007);
            if(!writeAndStartFile("bCSV.xlsx", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete bCSV.xlsx");

            var dta = new DataTable("andrei");
            dta.Columns.Add(new DataColumn("ID", typeof(int)));
            dta.Columns.Add(new DataColumn("Data", typeof(string)));

            dta.Rows.Add(1, "test 1 ");
            dta.Rows.Add(2, "test 2 ");
            dta.Rows.Add(3, "test 3 ");

            data = ExportFactory.ExportDataFromDataTable(dta, ExportToFormat.Excel2007);
            if(!writeAndStartFile("dta.xlsx", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete a.xlsx");

        }
        static void Advanced_ModifyTemplates()
        {
            var p = new Person { Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/", CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };
            var p1 = new Person { Name = "you", WebSite = "http://yourwebsite.com/" };
            var list = new List<Person>() { p, p1 };

            //advanced - modifying templates - in this case , added Number
            var export = new ExportExcel2007<Person>();
            //File.WriteAllText("Excel2007Header.txt",export.ExportHeader);
            //File.WriteAllText("Excel2007Item.txt", export.ExportItem);
            //File.WriteAllText("Excel2007Collection.txt", export.ExportCollection);
            export.ExportHeader = File.ReadAllText("Excel2007Header.txt");
            export.ExportItem = File.ReadAllText("Excel2007Item.txt");
            export.ExportCollection = File.ReadAllText("Excel2007Collection.txt");
            var data = export.ExportResult(list);
            if(!writeAndStartFile("WithId.xlsx", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete WithId.xlsx");
        }
    }
}
