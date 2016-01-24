using System;
using System.Configuration;
using ExportImplementation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExporterTests
{

    public partial class Excel2003Tests
    {
        [TestClass]
        public class TestConstructor
        {
            [TestMethod]
            public void TestConstructorHeaderWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var export = new ExportExcel2003<Person>();
                Assert.AreEqual(@"<Row>
   <Cell ss:StyleID='s21'><Data ss:Type='String'>Name</Data></Cell>
   <Cell ss:StyleID='s21'><Data ss:Type='String'>WebSite</Data></Cell>
   <Cell ss:StyleID='s21'><Data ss:Type='String'>CV</Data></Cell>
</Row>", export.ExportHeader);
            }

            [TestMethod]
            public void TestConstructorItemWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var export = new ExportExcel2003<Person>();
                Assert.AreEqual(@"<Row>
   <Cell><Data ss:Type='String'>@System.Security.SecurityElement.Escape((((object)Model.Name) ?? """").ToString())</Data></Cell>
   <Cell><Data ss:Type='String'>@System.Security.SecurityElement.Escape((((object)Model.WebSite) ?? """").ToString())</Data></Cell>
   <Cell><Data ss:Type='String'>@System.Security.SecurityElement.Escape((((object)Model.CV) ?? """").ToString())</Data></Cell>
</Row>".Replace("\r", "").Replace("\n", ""), export.ExportItem.Replace("\r", "").Replace("\n", ""));
            }
        }
    }
}
