using System;
using System.Configuration;
using ExportImplementation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExporterTests
{

    public partial class iTextSharp4Tests
    {
        [TestClass]
        public class TestConstructor
        {
            [TestMethod]
            public void TestConstructorHeaderWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var excel = new ExportPdfiTextSharp4<Person>();
                Assert.AreEqual(@"<row>
<cell><phrase font='Arial' size='12' style='bold'>Name</phrase></cell>
<cell><phrase font='Arial' size='12' style='bold'>WebSite</phrase></cell>
<cell><phrase font='Arial' size='12' style='bold'>CV</phrase></cell>
</row>", excel.ExportHeader);
            }

            [TestMethod]
            public void TestConstructorItemWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var excel = new ExportPdfiTextSharp4<Person>();
                Assert.AreEqual(@"<row>
    <cell><phrase font='Times New Roman' size='8'>@Model.Name</phrase></cell>
    <cell><phrase font='Times New Roman' size='8'>@Model.WebSite</phrase></cell>
    <cell><phrase font='Times New Roman' size='8'>@Model.CV</phrase></cell>
</row>", excel.ExportItem);
            }
        }
    }
}
