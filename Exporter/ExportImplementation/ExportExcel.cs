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
    public class ExportExcel2003<T> : Export<T>
        where T : class
    {
        public ExportExcel2003()
        {
            ExportCollection = Templates.Excel2003File;
            string template = Templates.Excel2003Header;
            var props = properties.Select(it => it.Name).ToArray();
            ExportHeader = Engine.Razor.RunCompile(template,TType.Name + "Excel2003HeaderInterpreter" + template.GetHashCode(),  typeof (string[]),props);
            template = Templates.Excel2003Item;
            ExportItem = Engine.Razor.RunCompile(template, TType.Name + "Excel2003ItemInterpreter" + template.GetHashCode(), typeof(string[]), props);

        }

        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var modelTemplate = new ModelTemplate<T>(data);
            
            var service = Engine.Razor;
            service.AddTemplate(TType.Name + "Excel2003Collection",ExportCollection);
            service.AddTemplate(TType.Name + "Excel2003Header", ExportHeader);
            service.AddTemplate(TType.Name + "Excel2003Item", ExportItem);
            service.Compile(TType.Name + "Excel2003Collection",typeof(ModelTemplate<T>));
            service.Compile(TType.Name + "Excel2003Header");
            service.Compile(TType.Name + "Excel2003Item",typeof(T));
            var result = service.Run(TType.Name + "Excel2003Collection", typeof(ModelTemplate<T>), modelTemplate);
            
            return System.Text.Encoding.Unicode.GetBytes(result);


        }

    }
}