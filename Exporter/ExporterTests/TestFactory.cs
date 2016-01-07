using System;
using System.Text;
using System.Collections.Generic;
using System.IO;
using ExportImplementation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExporterTests
{

    public class TestFactory
    {

        [TestClass]
        public class ExportDataJson
        {
            [TestMethod]
            public void TestCorrectJson()
            {
                string data = @"[
    { 'Name':'Andrei Ignat', 
        'WebSite':'http://msprogrammer.serviciipeweb.ro/',
        'CV':'http://serviciipeweb.ro/iafblog/content/binary/cv.doc'        
    },
{ 'Name':'Your Name', 
        'WebSite':'http://your website',
        'CV':'cv.doc'        
    }
]";
                

                var p = new Person { Name = "Andrei Ignat", WebSite = "http://msprogrammer.serviciipeweb.ro/", CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };
                var byte1 = new ExportExcel2007<Person>().ExportResult(new List<Person>(){p,p});

                var byte2 = ExportFactory.ExportDataJson(data, ExportToFormat.Excel2007);

                //File.WriteAllBytes("byte1.xlsx",byte1);
                //File.WriteAllBytes("byte2.xlsx", byte2);
                Assert.IsTrue(Math.Abs(byte1.Length-byte2.Length)<100);


            }

            [TestMethod]
            public void TestCorrectCSV()
            {
                List<string> data = new List<string>();
                data.Add("Name,WebSite,CV");
                data.Add("Andrei Ignat,http://msprogrammer.serviciipeweb.ro/,http://serviciipeweb.ro/iafblog/content/binary/cv.doc");
                data.Add("Andrei Ignat,http://msprogrammer.serviciipeweb.ro/,http://serviciipeweb.ro/iafblog/content/binary/cv.doc");

                var p = new Person { Name = "Andrei Ignat", WebSite = "http://msprogrammer.serviciipeweb.ro/", CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };

                var byte1 = new ExportExcel2007<Person>().ExportResult(new List<Person>() { p, p });

                var byte2 = ExportFactory.ExportDataCsv(data.ToArray(), ExportToFormat.Excel2007);

                Assert.IsTrue(Math.Abs(byte1.Length - byte2.Length) < 100);


            }
        }
    }
}