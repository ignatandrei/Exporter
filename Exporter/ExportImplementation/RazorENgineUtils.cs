using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RazorEngine.Templating;

namespace ExportImplementation
{
    static class RazorENgineUtils
    {
        /// <summary>
        /// adds CreatedDate if not exists
        /// </summary>
        /// <param name="additionalData"></param>
        public static DynamicViewBag ToDynamicViewBag(this KeyValuePair<string, object>[] additionalData)
        {
            var dict = new Dictionary<string, object> {{ "DateCreated", DateTime.Now}};
            if (additionalData != null && additionalData.Length > 0)
            {
                foreach (var item in additionalData)
                {
                    if (dict.ContainsKey(item.Key))
                    {
                        dict[item.Key] = item.Value;
                    }
                    else
                    {
                        dict.Add(item.Key, item.Value);
                    }
                }

            }
            var viewBag = new DynamicViewBag(dict);
            return viewBag;
        }
    }
}
