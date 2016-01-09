using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Channels;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ExporterObjects;
using Microsoft.SqlServer.Server;
using RazorEngine;
using RazorEngine.Templating;

namespace ExportImplementation
{
    public class ExportExcel2007<T> : Export<T>
        where T : class
    {
        public ExportExcel2007()
        {
            ExportCollection = Templates.Excel2007File;
            string template = Templates.Excel2007Header;
            var props = properties.Select(it => it.Name).ToArray();
            ExportHeader = Engine.Razor.RunCompile(template,TType.Name + "Excel2007HeaderInterpreter" + template.GetHashCode(),  typeof (string[]),props);
            template = Templates.Excel2007Item;
            ExportItem = Engine.Razor.RunCompile(template, TType.Name + "Excel2007ItemInterpreter" + template.GetHashCode(), typeof(string[]), props);

        }

        public string ExportResultStringPart(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            
            var modelTemplate = new ModelTemplate<T>(data);
            
            var service = Engine.Razor;
            service.AddTemplate(TType.Name + "Excel2007Collection",ExportCollection);
            service.AddTemplate(TType.Name + "Excel2007Header", ExportHeader);
            service.AddTemplate(TType.Name + "Excel2007Item", ExportItem);
            service.Compile(TType.Name + "Excel2007Collection",typeof(ModelTemplate<T>));
            service.Compile(TType.Name + "Excel2007Header");
            service.Compile(TType.Name + "Excel2007Item",typeof(T));
            var result = service.Run(TType.Name + "Excel2007Collection", typeof(ModelTemplate<T>), modelTemplate);
            
            return result;


        }
        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var result = ExportResultStringPart(data, additionalData);
            return CreateExcel2007(result);


        }
        private byte[] CreateExcel2007(string text)
        {
            using (var ms = new MemoryStream())
            {
                using (var sd = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbook = sd.AddWorkbookPart();
                    WorksheetPart sheet = workbook.AddNewPart<WorksheetPart>();
                    WriteToPart(sheet, text);
                    //TODO :put into a string 
                    WriteToPart(workbook, string.Format(
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"{1}\" sheetId=\"1\" r:id=\"{0}\" /></sheets></workbook>",
                        workbook.GetIdOfPart(sheet), this.TType.Name));

                    sd.Close();
                }
                return ms.ToArray();
            }
        }
        
        /// <summary>
        /// TODO: move into utilities
        /// </summary>
        /// <param name="oxp"></param>
        /// <param name="Text"></param>
        internal void WriteToPart(OpenXmlPart oxp, string Text)
        {
            using (Stream stream = oxp.GetStream())
            {
                byte[] buf = (new UTF8Encoding()).GetBytes(Text);
                stream.Write(buf, 0, buf.Length);
            }
        }

    }
}