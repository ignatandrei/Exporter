using System;
using System.Collections.Specialized;
using System.Runtime.Caching;
using System.Web;


namespace Swashbuckle.MVC
{
    public class SwashbuckleMVCModule:IHttpModule
    {
        public void Init(HttpApplication context)
        {
            context.BeginRequest += (o, e) =>
            {
                var localPath = context.Request.Url.LocalPath;
                if (localPath == "/swagger" || localPath == "/swagger/ui/index")
                {
                    context.Response.BufferOutput = true;
                    context.Response.Buffer = true;
                    var _watcher = new ResponseFilterStream(context.Response.Filter);                    
                    _watcher.TransformStream += _watcher_TransformStream;
                    context.Response.Filter = _watcher;
                }
            };
            

            //context.EndRequest += (o, e) =>
            context.PostRequestHandlerExecute += (o, e) =>
            {
                var localPath = context.Request.Url.LocalPath;
                context.Response.Flush();                
                
                if ((localPath == "/swagger" || localPath == "/swagger/ui/index")  )
                {
                    

                    var data = new Tuple<string, string>(Guid.NewGuid().ToString(), "ASDASD"+i);                    
                    var headers = new NameValueCollection();
                    headers["Guid"] = data.Item1;
                    MemoryCache.Default.Add(data.Item1, data, DateTimeOffset.Now.AddSeconds(10));
                    context.Server.TransferRequest("/SwashbuckleMVC",true,"GET",headers);
                    
                }

            };
        }

        private int i=0;
        private System.IO.MemoryStream _watcher_TransformStream(System.IO.MemoryStream arg)
        {
            i++;
            return arg;
        }

        

        public void Dispose()
        {
            
        }
    }
}
