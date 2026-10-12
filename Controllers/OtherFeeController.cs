using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;
using System.Configuration;
using WebApplication1.Models;
using WebApplication1.Services;

namespace WebApplication1.Controllers
{
    public class OtherFeeController : ApiController
    {
        private readonly OtherFeeService _service = new OtherFeeService();
        private readonly OtherFeeImportService _import = new OtherFeeImportService();
        private readonly OtherFeeExportService _export = new OtherFeeExportService();

        [HttpPost]
        public IHttpActionResult Add([FromBody] OtherFee fee)
        {
            if (fee == null) return BadRequest("请求体不能为空");
            var result = _service.Add(fee);
            if (result.errors.Count > 0)
                return Content(System.Net.HttpStatusCode.BadRequest, new { errors = result.errors });
            return Ok(new { id = result.id, message = "新增成功" });
        }

        [HttpGet]
        public IHttpActionResult List(string projectId)
        {
            if (string.IsNullOrWhiteSpace(projectId))
                return BadRequest("projectId 不能为空");
            var list = _service.List(projectId);
            return Ok(list);
        }

        [HttpPut]
        public IHttpActionResult Update(long id, [FromBody] OtherFee fee)
        {
            if (fee == null) return BadRequest("请求体不能为空");
            var errors = _service.Update(id, fee);
            if (errors.Count > 0)
                return Content(System.Net.HttpStatusCode.BadRequest, new { errors = errors });
            return Ok(new { message = "修改成功" });
        }

        [HttpDelete]
        public IHttpActionResult Delete(long id)
        {
            if (!_service.Delete(id))
                return NotFound();
            return Ok(new { message = "删除成功" });
        }

        [HttpPost, Route("api/otherfee/import")]
        public async Task<IHttpActionResult> Import(string projectId)
        {
            if (!Request.Content.IsMimeMultipartContent())
                return BadRequest("请用 multipart/form-data 上传文件");

            var provider = new MultipartMemoryStreamProvider();
            await Request.Content.ReadAsMultipartAsync(provider);

            var file = provider.Contents.FirstOrDefault();
            if (file == null) return BadRequest("未收到文件");

            using (var stream = await file.ReadAsStreamAsync())
            {
                var report = _import.Import(stream, projectId);
                return Ok(report);
            }
        }

        /// <summary>导出：GET api/otherfee/export?projectId=xxx&projectName=yyy → 下载 xlsx</summary>
        [HttpGet, Route("api/otherfee/export")]
        public IHttpActionResult Export(string projectId, string projectName = null)
        {
            if (string.IsNullOrWhiteSpace(projectId))
                return BadRequest("projectId 不能为空");
            if (string.IsNullOrWhiteSpace(projectName))
                projectName = projectId;

            byte[] bytes = _export.Export(projectId, projectName);

            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(bytes)
            };
            response.Content.Headers.ContentType =
                new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.Content.Headers.ContentDisposition =
                new ContentDispositionHeaderValue("attachment")
                {
                    FileName = "其他费用清单.xlsx"
                };
            return ResponseMessage(response);
        }

        /// <summary>重置：POST api/otherfee/reset（清空表重建）</summary>
        [HttpPost, Route("api/otherfee/reset")]
        public IHttpActionResult Reset(string key)
        {
            var expectedKey = ConfigurationManager.AppSettings["ResetKey"] ?? Repositories.Db.ResetKey;
            if (key != expectedKey)
                return BadRequest("无效的重置密钥");
            WebApplication1.Repositories.Db.Reset();
            return Ok(new { message = "数据库已重置" });
        }
    }
}
