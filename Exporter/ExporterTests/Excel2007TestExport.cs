using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Text;
using ExportImplementation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExporterTests
{

    public partial class Excel2007Tests
    {
        [TestClass]
        public class TestExport
        {
            [TestMethod]
            public void TestWithPersonHeader()
            {
                var p = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var export = new ExportExcel2007<Person>();
                var data = export.ExportResultStringPart(new List<Person>() {p});
                var str = data;
                Assert.IsTrue(str.Contains(export.ExportHeader),"must contain the header");
            }
            [TestMethod]
            public void TestWithPersonData()
            {
                var p = new Person { Name = "Andrei Ignat", WebSite = "http://msprogrammer.serviciipeweb.ro/", CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };
                var export = new ExportExcel2007<Person>();
                var data = export.ExportResultStringPart(new List<Person>() { p});
                var str = data;
                Assert.IsTrue(str.Contains("http://serviciipeweb.ro/iafblog/content/binary/cv.doc"),"must contain the cv");

            }
            [TestMethod]
            public void TestExcel2007MultipleSheetEqualJustOneSheet()
            {
                var p = new Person { Name = "Andrei Ignat", WebSite = "http://msprogrammer.serviciipeweb.ro/", CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };
                var list = new List<Person>() { p };
                var export = new ExportExcel2007<Person>();
                var data = export.ExportResult(list);

                var data1 = export.ExportMultipleSheets(new IList[] { list });
                Assert.IsTrue(Math.Abs(data.Length - data1.Length)<100);

            }
            //[TestMethod]
            public void TestExcel2007()
            {
                var p = new Person { Name = "Andrei Ignat", WebSite = "http://msprogrammer.serviciipeweb.ro/", CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };
                var export = new ExportExcel2007<Person>();
                var data = export.ExportResult(new List<Person>() { p });
                //var str = Encoding.Unicode.GetString(data);
                File.WriteAllBytes("a.xlsx",data);
                Process.Start("a.xlsx");

            }
            //[TestMethod]
            public void TestExcel2007MultipleSheets()
            {
                var p = new Person { Name = "Andrei Ignat", WebSite = "http://msprogrammer.serviciipeweb.ro/", CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };
                var list = new List<Person>() { p };
                var kvp = new List<Tuple<string, string>>();
                for (int i = 0; i < 10; i++)
                {
                    var q = new Tuple<string,string>("This is key " + i, "Value " + i);
                    kvp.Add(q);
                }
                
                var export = new ExportExcel2007<Person>();
                var data = export.ExportMultipleSheets(new IList[] {list, kvp});

                
                File.WriteAllBytes("multiple.xlsx", data);
                Process.Start("multiple.xlsx");

            }
        }
    }
}