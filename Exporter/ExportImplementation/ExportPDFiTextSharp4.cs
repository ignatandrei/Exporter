using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Channels;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ExporterObjects;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;
using Microsoft.SqlServer.Server;
using RazorEngine;
using RazorEngine.Templating;

namespace ExportImplementation
{
    public class ExportPdfiTextSharp4<T> : Export<T>
        where T : class
    {
        public ExportPdfiTextSharp4()
        {
            ExportCollection = Templates.iTextSharp4File;
            string template = Templates.iTextSharp4Header;
            var props = properties.Select(it => it.Name).ToArray();
            ExportHeader = Engine.Razor.RunCompile(template,TType.Name + "iTextSharp4HeaderInterpreter" + template.GetHashCode(),  typeof (string[]),props);
            template = Templates.iTextSharp4Item;
            ExportItem = Engine.Razor.RunCompile(template, TType.Name + "iTextSharp4ItemInterpreter" + template.GetHashCode(), typeof(string[]), props);

        }

        public string ExportResultStringPart(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var modelTemplate = new ModelTemplate<T>(data);
            
            var service = Engine.Razor;
            service.AddTemplate(TType.Name + "iTextSharp4Collection",ExportCollection);
            service.AddTemplate(TType.Name + "iTextSharp4Header", ExportHeader);
            service.AddTemplate(TType.Name + "iTextSharp4Item", ExportItem);
            service.Compile(TType.Name + "iTextSharp4Collection",typeof(ModelTemplate<T>));
            service.Compile(TType.Name + "iTextSharp4Header");
            service.Compile(TType.Name + "iTextSharp4Item",typeof(T));
            var result = service.Run(TType.Name + "iTextSharp4Collection", typeof(ModelTemplate<T>), modelTemplate, additionalData.ToDynamicViewBag());
            
            return result;


        }
        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var result = ExportResultStringPart(data, additionalData);
            return CreateiTextSharp4(result);


        }

        private byte[] CreateiTextSharp4(string result)
        {
            var xd=new XmlDocument();
            xd.LoadXml(result);
            using (var ms = new MemoryStream())
            {
                var document = new Document();
                PdfWriter.GetInstance(document, ms);
                XmlParser.Parse(document,xd);
                document.Close();
                return ms.ToArray();
            }


        }
    }
}