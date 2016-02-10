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
            
            return View();
        }
    }
}