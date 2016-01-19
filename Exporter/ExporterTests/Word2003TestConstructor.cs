using System;
using System.Configuration;
using ExportImplementation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExporterTests
{

    public partial class Word2003Tests
    {
        [TestClass]
        public class TestConstructor
        {
            [TestMethod]
            public void TestConstructorHeaderWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var export = new ExportWord2003<Person>();
                Assert.AreEqual(@"<w:tr>
           <w:tc>
                <w:p>
                    <w:r>
                        <w:rPr>
                            <w:b w:val='on'/>
                            <w:t>
                                Name
                            </w:t>
                        </w:rPr>
                    </w:r>
                </w:p>
            </w:tc>                
           <w:tc>
                <w:p>
                    <w:r>
                        <w:rPr>
                            <w:b w:val='on'/>
                            <w:t>
                                WebSite
                            </w:t>
                        </w:rPr>
                    </w:r>
                </w:p>
            </w:tc>                
           <w:tc>
                <w:p>
                    <w:r>
                        <w:rPr>
                            <w:b w:val='on'/>
                            <w:t>
                                CV
                            </w:t>
                        </w:rPr>
                    </w:r>
                </w:p>
            </w:tc>                
</w:tr>", export.ExportHeader);
            }

            [TestMethod]
            public void TestConstructorItemWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var export = new ExportWord2003<Person>();
                Assert.AreEqual(@"<w:tr>
   <w:tc>
    <w:p>
        <w:r>
            <w:t>@System.Security.SecurityElement.Escape((((object)Model.Name) ?? """").ToString())</w:t>
        </w:r>
    </w:p>
    </w:tc>
   <w:tc>
    <w:p>
        <w:r>
            <w:t>@System.Security.SecurityElement.Escape((((object)Model.WebSite) ?? """").ToString())</w:t>
        </w:r>
    </w:p>
    </w:tc>
   <w:tc>
    <w:p>
        <w:r>
            <w:t>@System.Security.SecurityElement.Escape((((object)Model.CV) ?? """").ToString())</w:t>
        </w:r>
    </w:p>
    </w:tc>
</w:tr>", export.ExportItem);
            }
        }
    }
}
