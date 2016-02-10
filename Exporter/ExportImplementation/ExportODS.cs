using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
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
    public class ExportODS<T> : Export<T>
        where T : class
    {
        public ExportODS()
        {
            ExportCollection = Templates.ODSFile;
            string template = Templates.ODSHeader;
            var props = properties.Select(it => it.Name).ToArray();
            ExportHeader = Engine.Razor.RunCompile(template,TType.Name + "ODSHeaderInterpreter" + template.GetHashCode(),  typeof (string[]),props);
            template = Templates.ODSItem;
            ExportItem = Engine.Razor.RunCompile(template, TType.Name + "ODSItemInterpreter" + template.GetHashCode(), typeof(string[]), props);

        }

        public string ExportResultStringPartNotGeneric(IList data, params KeyValuePair<string, object>[] additionalData)
        {

            var type = data[0].GetType();            
            dynamic exportType = typeof(ExportODS<>).MakeGenericType(type);
            dynamic export = Activator.CreateInstance(exportType);
            dynamic data1 = data;
            return export.ExportResultStringPart(data1, additionalData);

        }

        public string ExportResultStringPart(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            
            var modelTemplate = new ModelTemplate<T>(data);
            
            var service = Engine.Razor;
            service.AddTemplate(TType.Name + "ODSCollection",ExportCollection);
            service.AddTemplate(TType.Name + "ODSHeader", ExportHeader);
            service.AddTemplate(TType.Name + "ODSItem", ExportItem);
            service.Compile(TType.Name + "ODSCollection",typeof(ModelTemplate<T>));
            service.Compile(TType.Name + "ODSHeader");
            service.Compile(TType.Name + "ODSItem",typeof(T));
            var result = service.Run(TType.Name + "ODSCollection", typeof(ModelTemplate<T>), modelTemplate, additionalData.ToDynamicViewBag());
            
            return result;


        }
        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var result = ExportResultStringPart(data, additionalData);
            return CreateODS(result);


        }

     
        
        private byte[] CreateODS(string textSheet)
        {

            using (var ms = new MemoryStream())
            {
                ms.Write(Templates.ods, 0, Templates.ods.Length);
                using (var za = new ZipArchive(ms, ZipArchiveMode.Update))
                {
                    za.GetEntry("content.xml").Delete();
                    var c = za.CreateEntry("content.xml");
                    using (var s = c.Open())
                    {
                        using (var writer = new StreamWriter(s))
                        {
                            writer.Write(textSheet);
                        }
                    }
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