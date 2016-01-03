using System;
using System.Configuration;
using ExportImplementation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExporterTests
{

    public partial class HtmlTests
    {
        [TestClass]
        public class TestConstructor
        {
            [TestMethod]
            public void TestConstructorHeaderWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var excel = new ExportHtml<Person>();
                Assert.AreEqual(@"<tr>
   <th>Name</th>
   <th>WebSite</th>
   <th>CV</th>
</tr>", excel.ExportHeader);
            }

            [TestMethod]
            public void TestConstructorItemWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var excel = new ExportHtml<Person>();
                Assert.AreEqual(@"<tr>
   <td>@Model.Name</td>
   <td>@Model.WebSite</td>
   <td>@Model.CV</td>
</tr>", excel.ExportItem);
            }
        }
    }
}
