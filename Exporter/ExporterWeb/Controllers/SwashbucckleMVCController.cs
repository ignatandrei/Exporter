using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Caching;
using System.Web;
using System.Web.Mvc;

namespace ExporterWeb.Controllers
{
    public class SwashbuckleMVCController : Controller
    {
        // GET: SwashbucckleMVC
        public ActionResult Index()
        {
            var g = Request.Headers["Guid"];
            var t = MemoryCache.Default[g] as Tuple<string, string>;

            //t.Item2 = t.Item2.Replace("<html>", "").Replace("</html>", "").Replace("<head>", "").Replace("</head>", "");

            var t1 = new Tuple<string, string>(t.Item1, t.Item2.Replace("<!DOCTYPE html>","").Replace("<html>", "").Replace("</html>", "").Replace("<head>", "").Replace("</head>", ""));
            return View(t1);
        }
    }
}