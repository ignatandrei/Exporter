using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExporterObjects
{
    public abstract class Export<T>
    {
        protected PropertyInfo[] properties;
        protected Type TType;
        public Export()
        {
            TType= typeof (T);
            properties = TType.GetProperties();
            
        }
        public string ExportCollection { get; set; }
        public string ExportItem { get; set; }
        public string ExportHeader { get; set; }

        public abstract byte[] ExportResult(List<T> data,params KeyValuePair<string, object>[] additionalData);

    }
}