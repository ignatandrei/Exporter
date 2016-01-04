using System;
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
    public class ExportWord2007<T> : Export<T>
        where T : class
    {
        public ExportWord2007()
        {
            ExportCollection = Templates.Word2007File;
            string template = Templates.Word2007Header;
            var props = properties.Select(it => it.Name).ToArray();
            ExportHeader = Engine.Razor.RunCompile(template,TType.Name + "Word2007HeaderInterpreter",  typeof (string[]),props);
            template = Templates.Word2007Item;
            ExportItem = Engine.Razor.RunCompile(template, TType.Name + "Word2007ItemInterpreter", typeof(string[]), props);

        }

        public string ExportResultStringPart(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var modelTemplate = new ModelTemplate<T>(data);

            var service = Engine.Razor;
            service.AddTemplate(TType.Name + "Word2007Collection", ExportCollection);
            service.AddTemplate(TType.Name + "Word2007Header", ExportHeader);
            service.AddTemplate(TType.Name + "Word2007Item", ExportItem);
            service.Compile(TType.Name + "Word2007Collection", typeof(ModelTemplate<T>));
            service.Compile(TType.Name + "Word2007Header");
            service.Compile(TType.Name + "Word2007Item", typeof(T));
            var result = service.Run(TType.Name + "Word2007Collection", typeof(ModelTemplate<T>), modelTemplate);
            return result;
        }

        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var result = ExportResultStringPart(data, additionalData);
            return CreateWord2007(result);


        }
        private byte[] CreateWord2007(string Text)
        {
            using (var ms = new MemoryStream())
            {
                using (var wordDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    // Set the content of the document so that Word can open it.
                    var mainPart = wordDoc.AddMainDocumentPart();
                    WriteToPart(mainPart, Text);
                    wordDoc.Close();
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