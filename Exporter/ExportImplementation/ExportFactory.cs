using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExporterObjects;

namespace ExportImplementation
{
    public static class ExportFactory
    {
        public static byte[] ExportData<T>(List<T> data, ExportToFormat exportFormat, params KeyValuePair<string, object>[] additionalData)
            where T:class
        {
            Export<T> export;
            switch (exportFormat)
            {
                case ExportToFormat.Word2003XML:
                    export=new ExportWord2003<T>();
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
                    throw new ArgumentOutOfRangeException(nameof(exportFormat), exportFormat, null);
            }
            return export.ExportResult(data, additionalData);
        }

    }
}
