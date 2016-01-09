using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Caching;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Web.Http.Results;
using ExportImplementation;

namespace ExporterWeb.Controllers
{
    public class ExportData
    {
        public string data { get; set; }
    }
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class ExportController : ApiController
    {
        /// <summary>
        /// Export from CSV
        /// You should post a JSON.stringify({ 'data': data }) that has the CSV as string
        /// </summary>
        /// <param name="id"></param>
        /// <param name="dataCSV"></param>
        /// <returns>a GUID that is necessary to <see cref="GetFile"/> to obtain the file. Must be called in 10 minutes max</returns>
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
        /// <summary>
        /// Export from JSON
        /// You should post a JSON.stringify({ 'data': yourArray }) 
        /// </summary>
        /// <param name="id"></param>
        /// <param name="dataJson"></param>
        /// <returns>a GUID that is necessary to <see cref="GetFile"/> to obtain the file. Must be called in 10 minutes max</returns>
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

        /// <summary>
        /// Returns the file obtained by JSON/CSV
        /// </summary>
        /// <param name="id">the GUID obtained from <see cref="ExportFromCSV"/> or <see cref="ExportFromJSON"/></param>
        /// <param name="fileName">Name of the file for the user</param>
        /// <param name="extension">extension ( xlsx, docx, pdf...)</param>
        /// <returns>the file</returns>
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
