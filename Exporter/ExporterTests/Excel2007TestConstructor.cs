using System;
using System.Configuration;
using ExportImplementation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExporterTests
{

    public partial class Excel2007Tests
    {
        [TestClass]
        public class TestConstructor
        {
            [TestMethod]
            public void TestConstructorHeaderWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var export = new ExportExcel2007<Person>();
                Assert.AreEqual(@"<row>
   <c t='inlineStr'>
                <is>
                    <t>Name</t>
                </is>
            </c>
   <c t='inlineStr'>
                <is>
                    <t>WebSite</t>
                </is>
            </c>
   <c t='inlineStr'>
                <is>
                    <t>CV</t>
                </is>
            </c>
</row>", export.ExportHeader);
            }

            [TestMethod]
            public void TestConstructorItemWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var export = new ExportExcel2007<Person>();
                Assert.AreEqual(@"<row>
   <c t='inlineStr'>
                <is>
                    <t>@System.Security.SecurityElement.Escape((((object)Model.Name) ?? """").ToString())
                    </t>
                </is>
    </c>
   <c t='inlineStr'>
                <is>
                    <t>@System.Security.SecurityElement.Escape((((object)Model.WebSite) ?? """").ToString())
                    </t>
                </is>
    </c>
   <c t='inlineStr'>
                <is>
                    <t>@System.Security.SecurityElement.Escape((((object)Model.CV) ?? """").ToString())
                    </t>
                </is>
    </c>
</row>", export.ExportItem);
            }
        }
    }
}
