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
                var export = new ExportHtml<Person>();
                Assert.AreEqual(@"<tr>
   <th>Name</th>
   <th>WebSite</th>
   <th>CV</th>
</tr>", export.ExportHeader);
            }

            [TestMethod]
            public void TestConstructorItemWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var export = new ExportHtml<Person>();
                Assert.AreEqual(@"<tr>
   <td>@System.Security.SecurityElement.Escape((((object)Model.Name) ?? """").ToString())</td>
   <td>@System.Security.SecurityElement.Escape((((object)Model.WebSite) ?? """").ToString())</td>
   <td>@System.Security.SecurityElement.Escape((((object)Model.CV) ?? """").ToString())</td>
</tr>".Replace("\r","").Replace("\n", ""), export.ExportItem.Replace("\r", "").Replace("\n", ""));
            }
        }
    }
}
