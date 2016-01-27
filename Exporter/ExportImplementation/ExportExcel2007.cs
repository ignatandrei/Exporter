using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Channels;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
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

        public string ExportResultStringPartNotGeneric(IList data, params KeyValuePair<string, object>[] additionalData)
        {

            var type = data[0].GetType();            
            dynamic exportType = typeof(ExportExcel2007<>).MakeGenericType(type);
            dynamic export = Activator.CreateInstance(exportType);
            dynamic data1 = data;
            return export.ExportResultStringPart(data1, additionalData);

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
            var result = service.Run(TType.Name + "Excel2007Collection", typeof(ModelTemplate<T>), modelTemplate, additionalData.ToDynamicViewBag());
            
            return result;


        }
        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var result = ExportResultStringPart(data, additionalData);
            return CreateExcel2007(new string[] { TType.Name},new string[] { result});


        }

        public byte[] ExportMultipleSheets(IList[] data, params KeyValuePair<string, object>[] additionalData)
        {
            if (data == null)
                return null;
            if (data.Length == 0)
                return null;

            var result = new string[data.Length];
            var names = new string[data.Length];
            for (int i = 0; i < data.Length; i++)
            {
                var list = data[i];
                if(list == null || list.Count == 0)
                    continue;
                names[i] = list[0].GetType().Name;
                result[i] = ExportResultStringPartNotGeneric(list, additionalData);

            }
            return CreateExcel2007(names, result);




        }
        private byte[] CreateExcel2007(string[] worksheetName, string[] textSheet)
        {
            using (var ms = new MemoryStream())
            {
                using (var sd = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                {

                    var workbook = sd.AddWorkbookPart();
                    var strSheets = "<sheets>";
                    for (int i = 0; i < worksheetName.Length; i++)
                    {


                        var sheet = workbook.AddNewPart<WorksheetPart>();

                        WriteToPart(sheet, textSheet[i]);

                        strSheets += string.Format("<sheet name=\"{1}\" sheetId=\"{2}\" r:id=\"{0}\" />",
                            workbook.GetIdOfPart(sheet), worksheetName[i],(i+1));
                        
                        
                    }
                    strSheets += "</sheets>";
                    WriteToPart(workbook, string.Format(
                            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">{0}</workbook>",
                            strSheets
                            ));

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