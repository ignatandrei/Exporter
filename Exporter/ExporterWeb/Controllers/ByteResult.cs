using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace ExporterWeb.Controllers
{
    class ByteResult : IHttpActionResult
    {
        private readonly byte[] _data;
        private readonly string _fileName;
        private readonly string _contentType;

        public ByteResult(byte[] data,string fileName, string contentType = null)
        {
            

            _data= data;
            _fileName = fileName;
            _contentType = contentType;
        }

        public Task<HttpResponseMessage> ExecuteAsync(CancellationToken cancellationToken)
        {
            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(_data)
            };

            var contentType = _contentType ?? "application/octet-stream";
            response.Content.Headers.ContentType = new MediaTypeHeaderValue(contentType);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = _fileName
            };
            return Task.FromResult(response);
        }
    }
}