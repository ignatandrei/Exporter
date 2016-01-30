using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Net.Mime;
using System.Runtime.Caching;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;


namespace Swashbuckle.MVC
{
    public class SwashbuckleMVCModule:IHttpModule
    {
        class FakeController : ControllerBase { protected override void ExecuteCore() { } }
        public static string RenderViewToString(string controllerName, string viewName, object viewData)
        {
            using (var writer = new StringWriter())
            {
                var routeData = new RouteData();
                routeData.Values.Add("controller", controllerName);
                var fakeControllerContext = new ControllerContext(new HttpContextWrapper(new HttpContext(new HttpRequest(null, "http://google.com", null), new HttpResponse(null))), routeData, new FakeController());
                var razorViewEngine = new RazorViewEngine();
                var razorViewResult = razorViewEngine.FindView(fakeControllerContext, viewName, "", false);

                var viewContext = new ViewContext(fakeControllerContext, razorViewResult.View, new ViewDataDictionary(viewData), new TempDataDictionary(), writer);
                razorViewResult.View.Render(viewContext, writer);
                return writer.ToString();
            }
        }
        private static string Content;
        private static string ContentTop;
        private static string ContentDown;

        public void Init(HttpApplication context)
        {
            context.BeginRequest += (o, e) =>
            {
                var localPath = context.Request.Url.LocalPath;
                if (localPath == "/swagger" || localPath == "/swagger/ui/index")
                {
                    
                    if (Content == null)
                    {
                        Content = RenderViewToString("SwashbuckleMVC", "Index",
                            new Tuple<string, string>("!!!Content!!!", "!!!scripts!!!"));
                        string splitData = "!Here will be content!";
                        var index = Content.IndexOf(splitData);
                        ContentTop = Content.Substring(0, index);
                        ContentDown = Content.Substring(index + splitData.Length);
                    }
                    context.Response.BufferOutput = true;
                    context.Response.Buffer = true;
                    var id=Guid.NewGuid();
                    var _watcher = new ResponseFilterStream(context.Response.Filter,id);                    
                    _watcher.TransformStream += _watcher_TransformStream;
                    context.Response.Filter = _watcher;
                    page.Add(id,0);
                }
            };
            

            ////context.EndRequest += (o, e) =>
            //context.PostRequestHandlerExecute += (o, e) =>
            //{
            //    var localPath = context.Request.Url.LocalPath;
            //    context.Response.Flush();                
                
               

            //};
        }
        Dictionary<Guid,int> page=new Dictionary<Guid,int>();
        
        private System.IO.MemoryStream _watcher_TransformStream(Guid id,System.IO.MemoryStream ms)
        {
            var newMS = new MemoryStream();
            Encoding encoding = Encoding.Default;
            var val = page[id];
            page[id]=++val;
            if (val == 1) //first time
            {

                var output = encoding.GetBytes(ContentTop);                
                newMS.Write(output, 0, output.Length);
                
            }
            var arr = ms.ToArray();
            newMS.Write(arr, 0, arr.Length);
            //final sending of data - writing header
            if (ms.Length < Math.Min(4096,ms.Capacity))
            {
                var output = encoding.GetBytes(ContentDown);
                newMS.Write(output, 0, output.Length);
            }

            return newMS;
            return ms;
        }

        

        public void Dispose()
        {
            
        }
    }
}
