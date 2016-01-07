using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Caching;
using System.Web.Http;
using System.Web.Http.Results;
using ExportImplementation;

namespace ExporterWeb.Controllers
{
    public class ExportData
    {
        public string data { get; set; }
    }
    public class ExportController : ApiController
    {
        [HttpPost]
        public IHttpActionResult ExportFromCSV(
            ExportToFormat id, [FromBody]ExportData dataCSV)
        {
            try
            {
                var dataArray
                    = dataCSV.data.Trim()
                        .Split(new string[] {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries);

                if (dataArray.Length == 1)
                {
                    dataArray=dataCSV.data.Trim().Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                }

                var bytes = ExportFactory.ExportDataCsv(dataArray, id);
                var key = Guid.NewGuid().ToString();
                MemoryCache.Default.Add(key, bytes, DateTimeOffset.Now.AddMinutes(10));
                return Ok(key);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.StackTrace);
            }
        }

        [HttpPost]
        public IHttpActionResult ExportFromJSON([FromUri] ExportToFormat id, [FromBody] ExportData dataJson)
        {
            try
            {
                var bytes = ExportFactory.ExportDataJson(dataJson.data.Trim(), id);
                var key = Guid.NewGuid().ToString();
                MemoryCache.Default.Add(key, bytes, DateTimeOffset.Now.AddMinutes(10));
                return Ok(key);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.StackTrace);
            }
        }

        [HttpGet]
        public IHttpActionResult GetFile([FromUri] string id, [FromUri] string fileName, [FromUri] string extension)
        {
            try
            {
                if (!MemoryCache.Default.Contains(id))
                {
                    return this.NotFound();
                }
                var res = (byte[]) MemoryCache.Default[id];
                return new ByteResult(res, fileName + "." + extension);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.StackTrace);
            }
        }

    }
}
