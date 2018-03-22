using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace OfficeViewer.Controllers
{
    [Route("File")]
    [Authorize]
    public class FileController : Controller
    {
        private const String AccessKey = "123456";

        private IHostingEnvironment _env;

        public FileController(IHostingEnvironment env)
        {
            _env = env;
        }

        [Route("GetDocument/{key}/{file}")]
        [AllowAnonymous]
        public async Task<IActionResult> GetDocument([FromRoute]string file, string key)
        {
            if (AccessKey != key)
            {
                return Unauthorized();
            }
            var fileStream = Path.Combine(_env.ContentRootPath, "Files/" + file);
            FileStream stream = new FileStream(fileStream, FileMode.Open);
            Response.Headers.Add("accept-ranges", "bytes");
            return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        }

        [Route("ViewDocument")]
        public async Task<IActionResult> ViewDocument([FromQuery]string file)
        {
            var appPath = Request.Host.Value;
            var filePath = $"{appPath}/File/GetDocument/{AccessKey}/{file}";
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri("https://view.officeapps.live.com/");
                var data = await client.GetAsync("/op/view.aspx?src=" + filePath);
                var content = await data.Content.ReadAsByteArrayAsync();
                Response.Body.Write(content, 0, content.Length);
                Response.Body.Flush();
            }
            return Ok();
        }

    }
}