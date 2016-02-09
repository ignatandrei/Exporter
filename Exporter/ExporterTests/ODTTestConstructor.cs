using System;
using System.Configuration;
using ExportImplementation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExporterTests
{

    public partial class ODTTests
    {
        [TestClass] 
        public class TestConstructor
        {
            [TestMethod]
            public void TestConstructorHeaderWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var export = new ExportODT<Person>();
                Assert.AreEqual(@"<table:table-column table:style-name='Standard'  table:number-columns-repeated='3'/> 
<table:table-row> 
               <table:table-cell table:style-name='Standard'
		  office:value-type='string'>

		  	<text:p>
                                      Name
                              </text:p>
		
                </table:table-cell>
               <table:table-cell table:style-name='Standard'
		  office:value-type='string'>

		  	<text:p>
                                      WebSite
                              </text:p>
		
                </table:table-cell>
               <table:table-cell table:style-name='Standard'
		  office:value-type='string'>

		  	<text:p>
                                      CV
                              </text:p>
		
                </table:table-cell>
</table:table-row>".Replace("\r", "").Replace("\n", ""), export.ExportHeader.Replace("\r", "").Replace("\n", ""));
            }

            [TestMethod]
            public void TestConstructorItemWithPerson()
            {
                var t = new Person {Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/"};
                var export = new ExportODT<Person>();
                Assert.AreEqual(@"<table:table-row> 
<table:table-cell table:style-name='Standard'
 office:value-type='string'>
		  	
<text:p>@System.Security.SecurityElement.Escape((((object)Model.Name) ?? """").ToString())</text:p>

</table:table-cell>   
<table:table-cell table:style-name='Standard'
 office:value-type='string'>
		  	
<text:p>@System.Security.SecurityElement.Escape((((object)Model.WebSite) ?? """").ToString())</text:p>

</table:table-cell>   
<table:table-cell table:style-name='Standard'
 office:value-type='string'>
		  	
<text:p>@System.Security.SecurityElement.Escape((((object)Model.CV) ?? """").ToString())</text:p>

</table:table-cell>   
</table:table-row>".Replace("\r", "").Replace("\n", ""), export.ExportItem.Replace("\r", "").Replace("\n", ""));
            }
        }
    }
}
