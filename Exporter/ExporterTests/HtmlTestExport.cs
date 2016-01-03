using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Text;
using ExportImplementation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExporterTests
{

    public partial class HtmlTests
    {
        [TestClass]
        public class TestExport
        {
            [TestMethod]
            public void TestWithPersonHeader()
            {
                var p = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var excel = new ExportHtml<Person>();
                var data = excel.ExportResult(new List<Person>() {p});
                var str = Encoding.Unicode.GetString(data);
                Assert.IsTrue(str.Contains(excel.ExportHeader),"must contain the header");
            }
            [TestMethod]
            public void TestWithPersonData()
            {
                var p = new Person { Name = "Andrei Ignat", WebSite = "http://msprogrammer.serviciipeweb.ro/", CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };
                var excel = new ExportHtml<Person>();
                var data = excel.ExportResult(new List<Person>() { p});
                var str = Encoding.Unicode.GetString(data);
                Assert.IsTrue(str.Contains("http://serviciipeweb.ro/iafblog/content/binary/cv.doc"),"must contain the cv");

            }
            [TestMethod]
            public void TestHtml()
            {
                var p = new Person { Name = "Andrei Ignat", WebSite = "http://msprogrammer.serviciipeweb.ro/", CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };
                var excel = new ExportHtml<Person>();
                var data = excel.ExportResult(new List<Person>() { p });
                var str = Encoding.Unicode.GetString(data);
                File.WriteAllText("a.html",str);
                Process.Start("a.html");

            }
        }
    }
}