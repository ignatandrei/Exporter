using System;
using System.Collections;
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
            Console.WriteLine("3. See advanced - export multiple sheets");

            var option = Console.ReadLine();
            switch (option)
            {
                case "1":
                    SeeAllExport();
                    return;
                case "2":
                    Advanced_ModifyTemplates();
                    return;
                case "3":
                    Advanced_MultipleSheets();
                    return;
            }
            Console.WriteLine("demo finished");
            
           

        }

        private static void Advanced_MultipleSheets()
        {
            var p = new Person { Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/", CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };
            var p1 = new Person { Name = "you", WebSite = "http://yourwebsite.com/" };
            var list = new List<Person>() { p, p1 };

            var kvp = new List<Tuple<string, string>>();
            for (int i = 0; i < 10; i++)
            {
                var q = new Tuple<string, string>("This is key " + i, "Value " + i);
                kvp.Add(q);
            }

            var export = new ExportExcel2007<Person>();
            var data = export.ExportMultipleSheets(new IList[] { list, kvp });
            if (!writeAndStartFile("multiple.xlsx", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete multiple.xlsx");


            //export dataset
            var ds = new DataSet();
            var table = new DataTable("programmers");
            var idColumn = table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("WebSite", typeof(string));

            table.PrimaryKey = new DataColumn[] { idColumn };

            table.Rows.Add(new object[] { 1, "Andrei Ignat", "http://msprogrammer.serviciipeweb.ro" });
            table.Rows.Add(new object[] { 2, "Scott Hanselman", "http://www.hanselman.com/blog/" });

            ds.Tables.Add(table);

            var dta = new DataTable("andrei");
            dta.Columns.Add(new DataColumn("ID", typeof(int)));
            dta.Columns.Add(new DataColumn("Data", typeof(string)));
            dta.Rows.Add(1, "test 1 ");
            dta.Rows.Add(2, "test 2 ");
            dta.Rows.Add(3, "test 3 ");
            ds.Tables.Add(dta);



            export = new ExportExcel2007<Person>();
            data = ExportFactory.ExportDataSet(ds, ExportToFormat.Excel2007);
            if (!writeAndStartFile("multipleDataSet.xlsx", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete multipleDataSet.xlsx");
            
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

            export = new ExportExcel2003<Person>();
            data = export.ExportResult(list);            
            if (!writeAndStartFile("a.xls", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete a.xls");

            export = new ExportODS<Person>();
            data = export.ExportResult(list);
            if (!writeAndStartFile("a.ods", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete a.ods");


            export = new ExportExcel2007<Person>();
            data = export.ExportResult(list);
            if(!writeAndStartFile("a.xlsx", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete a.xlsx");

            export = new ExportWord2003<Person>();
            data = export.ExportResult(list);
            if (!writeAndStartFile("a.doc", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete a.doc");



            export = new ExportWord2007<Person>();
            data = export.ExportResult(list);
            if(!writeAndStartFile("a.docx", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete a.docx");


            export = new ExportPdfiTextSharp4<Person>();
            data = export.ExportResult(list);
            if(!writeAndStartFile("a.pdf", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete a.pdf");


            export = new ExportODT<Person>();
            data = export.ExportResult(list);
            if (!writeAndStartFile("a.odt", data))
                Console.WriteLine(" !!!!!!!!!!Could not delete a.odt");


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
