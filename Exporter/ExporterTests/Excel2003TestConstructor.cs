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
            public void TestConstructorWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var excel = new ExportExcel<Person>();
                Assert.AreEqual(@"<Row>
   <Cell><Data ss:Type='String'>Name</Data></Cell>
   <Cell><Data ss:Type='String'>WebSite</Data></Cell>
</Row>",excel.ExportHeader);
            }
        }
    }
}
