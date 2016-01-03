using System;
using System.Collections.Generic;

namespace ExportImplementation
{
    public class ModelTemplate<T>
    {
        protected Type TType;
        public ModelTemplate(List<T> data )
        {
            TType = typeof(T);
            NameOfT = TType.Name;
            Data = data;
        } 
        public string NameOfT { get; set; }
        public List<T> Data { get; set; }


    }
}