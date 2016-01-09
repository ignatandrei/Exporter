using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ExporterObjects;
using Microsoft.CSharp;
using Newtonsoft.Json.Linq;
using RazorEngine;
using RazorEngine.Templating;

namespace ExportImplementation
{
    public static class ExportFactory
    {
        public static byte[] ExportDataWithType(IEnumerable data, ExportToFormat exportFormat,Type type,
            params KeyValuePair<string, object>[] additionalData)
            
        {
            var exportType = typeof(Export<>).MakeGenericType(type);

            switch (exportFormat)
            {
                case ExportToFormat.Word2003XML:
                    
                    exportType = typeof(ExportWord2003<>).MakeGenericType(type);
                    break;
                case ExportToFormat.Excel2003XML:
                    exportType = typeof(ExportExcel2003<>).MakeGenericType(type);                    
                    break;
                case ExportToFormat.HTML:
                    exportType = typeof(ExportHtml<>).MakeGenericType(type);
                    break;
                case ExportToFormat.PDFiTextSharp4:
                    exportType = typeof(ExportPdfiTextSharp4<>).MakeGenericType(type);
                    break;
                case ExportToFormat.Word2007:
                    exportType = typeof(ExportWord2007<>).MakeGenericType(type);
                    break;
                case ExportToFormat.Excel2007:
                    exportType = typeof(ExportExcel2007<>).MakeGenericType(type);
                    break;
                default:
                    //throw new ArgumentOutOfRangeException(nameof(exportFormat), exportFormat, null);
                    throw new ArgumentOutOfRangeException("exportFormat", exportFormat, null);
            }

            dynamic export = Activator.CreateInstance(exportType);
            dynamic data1 = data;
            return export.ExportResult(data1, additionalData);

        }

        public static byte[] ExportData<T>(List<T> data, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
            where T : class
        {
            Export<T> export;
            switch (exportFormat)
            {
                case ExportToFormat.Word2003XML:
                    export = new ExportWord2003<T>();
                    break;
                case ExportToFormat.Excel2003XML:
                    export = new ExportExcel2003<T>();
                    break;
                case ExportToFormat.HTML:
                    export = new ExportHtml<T>();
                    break;
                case ExportToFormat.PDFiTextSharp4:
                    export = new ExportPdfiTextSharp4<T>();
                    break;
                case ExportToFormat.Word2007:
                    export = new ExportWord2007<T>();
                    break;
                case ExportToFormat.Excel2007:
                    export = new ExportExcel2007<T>();
                    break;
                default:
                    //throw new ArgumentOutOfRangeException(nameof(exportFormat), exportFormat, null);
                    throw new ArgumentOutOfRangeException("exportFormat", exportFormat, null);
            }
            return export.ExportResult(data, additionalData);
        }

        public class ModelRuntime
        {
            public string ClassName { get; set; }
            public string[] Properties { get; set; }
        }

        private static Type FromProperties(string[] props)
        {
            var constructor = string.Join(",string ", props);
            var hash = constructor.GetHashCode();

            var mrj = new ModelRuntime();
            mrj.ClassName = "Data" + hash;
            try
            {
                var typeExisting= AppDomain.CurrentDomain.GetAssemblies()            
                    .SelectMany(a => a.GetTypes())
                    .FirstOrDefault(t => t.FullName.Equals(mrj.ClassName));

                if (typeExisting != null)
                    return typeExisting;

            }
            catch (Exception ex)
            {
                string s = ex.Message;
            }
            mrj.Properties = props;

            var template = @"
using System;
public class @Model.ClassName {
//constructor
public @Model.ClassName (
    @foreach(var prop in Model.Properties){
    <text>string @prop , </text>
    }
    //add a fake property
string fake=null)
{
    @foreach(var prop in Model.Properties){
    <text>this.@prop = @prop;</text>
    }
}//end constructor


//properties
@foreach(var prop in Model.Properties){
    <text>public string @prop{get;set;}</text>
    }


 
}//end class               
";
            var code = Engine.Razor.RunCompile(template, mrj.ClassName, typeof(ModelRuntime), mrj);
            var provider = new CSharpCodeProvider();
            var parameters = new CompilerParameters();
            parameters.ReferencedAssemblies.Add("System.dll");
            // True - memory generation, false - external file generation
            parameters.GenerateInMemory = false;
            // True - exe file generation, false - dll file generation
            parameters.GenerateExecutable = false;

            var results = provider.CompileAssemblyFromSource(parameters, code);
            if (results.Errors.HasErrors)
            {
                var sb = new StringBuilder();

                foreach (CompilerError error in results.Errors)
                {
                    sb.AppendLine(string.Format("Error ({0}): {1}", error.ErrorNumber, error.ErrorText));
                }

                throw new InvalidOperationException(sb.ToString());
            }
            var assembly = results.CompiledAssembly;

            
            

            var type = assembly.DefinedTypes.First(t => t.Name == mrj.ClassName);
            return type;
        }
        public static byte[] ExportDataJson(string jsonArray, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {
            var jObj = JArray.Parse(jsonArray);
            var props = jObj[0].Select(it => ((JProperty) it).Name).ToArray();
            var type = FromProperties(props);
            var assembly = type.Assembly;
            //in order to be found from Razor export
            Assembly.LoadFile(assembly.Location);
            var listType = typeof(List<>).MakeGenericType(type);
            
            dynamic list = Activator.CreateInstance(listType);
            for (int i=0;i<jObj.Count;i++)
            {
                var item = jObj[i];
                var propsValue = item.Select(it => ((JProperty)it).Value).Select(it=>it.ToString()).ToList();
                propsValue.Add("fake");
                dynamic obj = assembly.CreateInstance(type.FullName, true, BindingFlags.Public | BindingFlags.Instance, null,
                    propsValue.ToArray(), null,
                    null);
                list.Add(obj);

            }
            return ExportDataWithType(list as IEnumerable, exportFormat, type, additionalData);
        }

        public static byte[] ExportDataCsv(string[] csvWithHeader, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {

            var props = csvWithHeader[0].Split(new string[] {","}, StringSplitOptions.None);
            if(props.Contains(""))
                throw new ArgumentException("header contains empty string");
            
            var type = FromProperties(props);
            var assembly = type.Assembly;
            //in order to be found from Razor export
            Assembly.LoadFile(assembly.Location);
            var listType = typeof(List<>).MakeGenericType(type);

            dynamic list = Activator.CreateInstance(listType);
            for (int i = 1; i < csvWithHeader.Length; i++)
            {
                var item = csvWithHeader[i];
                var propsValue = item.Split(new string[] {","}, StringSplitOptions.None).ToList();
                propsValue.Add("fake");
                dynamic obj = assembly.CreateInstance(type.FullName, true, BindingFlags.Public | BindingFlags.Instance, null,
                    propsValue.ToArray(), null,
                    null);
                list.Add(obj);

            }
            return ExportDataWithType(list as IEnumerable, exportFormat, type, additionalData);
        }


        public static byte[] ExportDataFromDataTable(DataTable data, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {
            var cols = data.Columns;

            var props = new string[cols.Count];
            for (int i = 0; i < cols.Count; i++)
            {
                props[i] = cols[i].ColumnName;
            }

            var type = FromProperties(props);
            var assembly = type.Assembly;
            //in order to be found from Razor export
            Assembly.LoadFile(assembly.Location);
            var listType = typeof(List<>).MakeGenericType(type);

            dynamic list = Activator.CreateInstance(listType);
            for (int i = 0; i < data.Rows.Count; i++)
            {
                var item = data.Rows[i];
                var propsValue = item.ItemArray.Select(it => it.ToString()).ToList();
                propsValue.Add("fake");
                dynamic obj = assembly.CreateInstance(type.FullName, true, BindingFlags.Public | BindingFlags.Instance, null,
                    propsValue.ToArray(), null,
                    null);
                list.Add(obj);

            }
            return ExportDataWithType(list as IEnumerable, exportFormat, type, additionalData);
        }
    }
}