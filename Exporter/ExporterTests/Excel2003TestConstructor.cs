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
                var excel = new ExportExcel<Person>();
                Assert.AreEqual(@"<Row>
   <Cell ss:StyleID='s21'><Data ss:Type='String'>Name</Data></Cell>
   <Cell ss:StyleID='s21'><Data ss:Type='String'>WebSite</Data></Cell>
   <Cell ss:StyleID='s21'><Data ss:Type='String'>CV</Data></Cell>
</Row>", excel.ExportHeader);
            }

            //[TestMethod]
            public void TestConstructorItemWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var excel = new ExportExcel<Person>();
                Assert.AreEqual(@"<Row>
   <Cell><Data ss:StyleID='s21' ss:Type='String'>@Model.Name</Data></Cell>
   < Cell><Data ss:StyleID='s21' ss:Type='String'>@Model.WebSite</Data></Cell>
   <Cell><Data ss:StyleID='s21' ss:Type='String'>@Model.CV</Data></Cell>
</Row>", excel.ExportItem);
            }
        }
    }
}
