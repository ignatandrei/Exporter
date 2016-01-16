using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Channels;
using System.Text;
using System.Threading.Tasks;
using ExporterObjects;
using Microsoft.SqlServer.Server;
using RazorEngine;
using RazorEngine.Templating;

namespace ExportImplementation
{
    public class ExportHtml<T> : Export<T>
        where T : class
    {
        public ExportHtml()
        {
            ExportCollection = Templates.HtmlFile;
            string template = Templates.HtmlHeader;
            var props = properties.Select(it => it.Name).ToArray();
            ExportHeader = Engine.Razor.RunCompile(template,TType.Name + "HtmlHeaderInterpreter" + template.GetHashCode(),  typeof (string[]),props);
            template = Templates.HtmlItem;
            ExportItem = Engine.Razor.RunCompile(template, TType.Name + "HtmlItemInterpreter" + template.GetHashCode(), typeof(string[]), props);

        }

        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var modelTemplate = new ModelTemplate<T>(data);
            
            var service = Engine.Razor;
            service.AddTemplate(TType.Name + "HtmlCollection",ExportCollection);
            service.AddTemplate(TType.Name + "HtmlHeader", ExportHeader);
            service.AddTemplate(TType.Name + "HtmlItem", ExportItem);
            service.Compile(TType.Name + "HtmlCollection",typeof(ModelTemplate<T>));
            service.Compile(TType.Name + "HtmlHeader");
            service.Compile(TType.Name + "HtmlItem",typeof(T));
            var result = service.Run(TType.Name + "HtmlCollection", typeof(ModelTemplate<T>), modelTemplate, additionalData.ToDynamicViewBag());
            
            return System.Text.Encoding.Unicode.GetBytes(result);


        }

    }
}