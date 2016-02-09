using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
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
using Encoding = System.Text.Encoding;

namespace ExportImplementation
{
    public class ExportODT<T> : Export<T>
        where T : class
    {
        public ExportODT()
        {
            ExportCollection = Templates.ODTFile;
            string template = Templates.ODTHeader;
            var props = properties.Select(it => it.Name).ToArray();
            ExportHeader = Engine.Razor.RunCompile(template,TType.Name + "ODTHeaderInterpreter",  typeof (string[]),props);
            template = Templates.ODTItem;
            ExportItem = Engine.Razor.RunCompile(template, TType.Name + "ODTItemInterpreter", typeof(string[]), props);

        }

        public string ExportResultStringPart(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var modelTemplate = new ModelTemplate<T>(data);

            var service = Engine.Razor;
            service.AddTemplate(TType.Name + "ODTCollection", ExportCollection);
            service.AddTemplate(TType.Name + "ODTHeader", ExportHeader);
            service.AddTemplate(TType.Name + "ODTItem", ExportItem);
            service.Compile(TType.Name + "ODTCollection", typeof(ModelTemplate<T>));
            service.Compile(TType.Name + "ODTHeader");
            service.Compile(TType.Name + "ODTItem", typeof(T));
            var result = service.Run(TType.Name + "ODTCollection", typeof(ModelTemplate<T>), modelTemplate, additionalData.ToDynamicViewBag());
            return result;
        }

        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var result = ExportResultStringPart(data, additionalData);
            return CreateODT(result);


        }
        private byte[] CreateODT(string Text)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(Templates.ODT,0, Templates.ODT.Length);
                using (var za = new ZipArchive(ms, ZipArchiveMode.Update))
                {
                    za.GetEntry("content.xml").Delete();
                    var c = za.CreateEntry("content.xml");
                    using (var s = c.Open())
                    {
                        using (var writer = new StreamWriter(s))
                        {
                            writer.Write(Text);
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
                byte[] buf =  (new UTF8Encoding()).GetBytes(Text);
                stream.Write(buf, 0, buf.Length);
            }
        }
    }
}