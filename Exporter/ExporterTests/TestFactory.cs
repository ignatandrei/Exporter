using System;
using System.Text;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using ExportImplementation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;

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
                var data = new List<string>();
                data.Add("Name,WebSite,CV");
                data.Add("Andrei Ignat,http://msprogrammer.serviciipeweb.ro/,http://serviciipeweb.ro/iafblog/content/binary/cv.doc");
                data.Add("Andrei Ignat,http://msprogrammer.serviciipeweb.ro/,http://serviciipeweb.ro/iafblog/content/binary/cv.doc");

                var p = new Person { Name = "Andrei Ignat", WebSite = "http://msprogrammer.serviciipeweb.ro/", CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };

                var byte1 = new ExportExcel2007<Person>().ExportResult(new List<Person>() { p, p });

                var byte2 = ExportFactory.ExportDataCsv(data.ToArray(), ExportToFormat.Excel2007);

                Assert.IsTrue(Math.Abs(byte1.Length - byte2.Length) < 100);


            }


            [TestMethod]
            public void TestCorrectDataTable()
            {
                var dta = new DataTable("andrei");
                dta.Columns.Add(new DataColumn("ID", typeof(int)));
                dta.Columns.Add(new DataColumn("Data", typeof(string)));
                dta.Rows.Add(1, "test 1 ");
                dta.Rows.Add(2, "test 2 ");
                dta.Rows.Add(3, "test 3 ");
                
                var byte1 = ExportFactory.ExportDataFromDataTable(dta, ExportToFormat.Excel2007);
                //File.WriteAllBytes("a.xlsx",byte1);
                //Process.Start("a.xlsx");
                Assert.IsTrue(Math.Abs(byte1.Length - 1921) < 100,byte1.Length.ToString());


            }

            [TestMethod]
            public async Task TestJsonDownloadWithCSharp()
            {
                var url = "https://api.github.com/search/users?q=ignata";
                var client = new HttpClient();
                client.DefaultRequestHeaders.Add("user-agent", "AA");
                var json = await client.GetStringAsync(url);
                var jObj = JObject.Parse(json);
                var data = ExportFactory.ExportDataJson(jObj["items"].ToString(), ExportToFormat.Excel2007);
                
                //File.WriteAllBytes("a.xlsx", data);
                //Process.Start("a.xlsx");
                Assert.IsTrue(Math.Abs(data.Length-2525) < 100, data.Length.ToString());
            }

            [TestMethod]
            public void TestCorrectRSS()
            {
                var data=ExportFactory.ExportDataRSS("http://msprogrammer.serviciipeweb.ro/feed/", ExportToFormat.Excel2003XML);
                var str=Encoding.UTF8.GetString(data);

                //File.WriteAllText("a.xls", str);
                //Process.Start("a.xls");
                Assert.IsTrue(str.Contains("msprogrammer.serviciipeweb.ro"), "must contain link to my site");
            }
            [TestMethod]
            public void TestIDataReader()
            {
                var table = new DataTable();
                var idColumn = table.Columns.Add("ID", typeof(int));
                table.Columns.Add("Name", typeof(string));
                table.Columns.Add("WebSite", typeof(string));

                table.PrimaryKey = new DataColumn[] { idColumn };

                table.Rows.Add(new object[] { 1, "Andrei Ignat", "http://msprogrammer.serviciipeweb.ro" });
                table.Rows.Add(new object[] { 2, "Scott Hanselman", "http://www.hanselman.com/blog/" });

                var byteDataTable = ExportFactory.ExportDataFromDataTable(table, ExportToFormat.Excel2003XML);
                var byteIDataReader = ExportFactory.ExportDataIDataReader(table.CreateDataReader(), ExportToFormat.Excel2003XML);
                Assert.AreEqual(byteDataTable.Length,byteIDataReader.Length);
                //var str = Encoding.Unicode.GetString(byteDataTable);
                //File.WriteAllText("a.xls", str);
                //Process.Start("a.xls");
            }

            [TestMethod]
            public void TestOpmlRSS()
            {
                string opml = @"<?xml version='1.0' encoding='utf-8'?>
<opml version='1.0'>
	<head>
		<title>My Blogs</title>
	</head>
	<body>
		<outline text='Programming Blog En' htmlUrl='http://msprogrammer.serviciipeweb.ro/' type='rss' xmlUrl='http://msprogrammer.serviciipeweb.ro/feed/'/>
		<outline text='Programming Blog RO' htmlUrl='http://serviciipeweb.ro/iafblog/' type='rss' xmlUrl='http://serviciipeweb.ro/iafblog/feed/'/>
		<outline text='Personal Blog RO' htmlUrl='http://serviciipeweb.ro/propriu/' type='rss' xmlUrl='http://serviciipeweb.ro/propriu/feed'/>

	</body>
</opml> ";

                var data = ExportFactory.ExportOpmlRSS(opml,ExportToFormat.Excel2007);
                //File.WriteAllBytes("opml.xlsx",data);
                //Process.Start("opml.xlsx");
                Assert.IsTrue(data.Length>1000 );
            }
        }
    }
}