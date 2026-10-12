using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using WebApplication1.Services;

namespace WebApplication1.Controllers
{
    public class CgroupController : ApiController
    {
        private readonly CgroupImportService _import = new CgroupImportService();
        private readonly BoomImportService _boomImport = new BoomImportService();
        private readonly Repositories.CgroupRepository _repo = new Repositories.CgroupRepository();

        // bin 目录下的两个输入模板
        private static string BinDir => AppDomain.CurrentDomain.BaseDirectory;
        private static string GongfaFile => Path.Combine(BinDir, "工法清单模板--2.定额计价工法导入.xlsx");
        private static string MainMatFile => Path.Combine(BinDir, "工法清单模板--1.主材导入.xlsx");

        /// <summary>触发导入：从 bin 读取两个模板，写入SQLite</summary>
        [HttpPost]
        [Route("api/cgroup/import")]
        public IHttpActionResult Import()
        {
            try
            {
                var r = _import.Import(GongfaFile, MainMatFile);
                var boomN = _boomImport.Import(Path.Combine(BinDir, "分部分项清单.json"));
                return Ok(new
                {
                    message = "C组导入完成",
                    gongfaCount = r.GongFaCount,
                    resourceCount = r.ResourceCount,
                    mainMaterialCount = r.MainMaterialCount,
                    feeRuleCount = r.FeeRuleCount,
                    boomCount = boomN
                });
            }
            catch (Exception ex)
            {
                return InternalServerError(ex);
            }
        }

        /// <summary>查看算钱结果（临时验证用）</summary>
        [HttpGet]
        [Route("api/cgroup/price")]
        public IHttpActionResult Price()
        {
            var calc = new CgroupCalcService();
            return Ok(calc.CalcAll());
        }

        /// <summary>导出工法主材汇总表</summary>
        [HttpGet]
        [Route("api/cgroup/export")]
        public IHttpActionResult Export()
        {
            var export = new CgroupExportService();
            byte[] bytes = export.ExportGongFa();

            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(bytes)
            };
            response.Content.Headers.ContentType =
                new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.Content.Headers.ContentDisposition =
                new ContentDispositionHeaderValue("attachment")
                {
                    FileName = "工法主材汇总表.xlsx"
                };
            return ResponseMessage(response);
        }

        /// <summary>导入分部分项 JSON</summary>
        [HttpPost]
        [Route("api/cgroup/boomimport")]
        public IHttpActionResult BoomImport()
        {
            string json = Path.Combine(BinDir, "分部分项清单.json");
            int n = _boomImport.Import(json);
            return Ok(new { message = "分部分项导入完成", count = n });
        }

        /// <summary>查看分部分项数据</summary>
        [HttpGet]
        [Route("api/cgroup/boom")]
        public IHttpActionResult Boom()
        {
            var repo = new Repositories.CgroupRepository();
            return Ok(repo.ListBoom());
        }

        /// <summary>导出分部分项清单</summary>
        [HttpGet]
        [Route("api/cgroup/boomexport")]
        public IHttpActionResult BoomExport()
        {
            var export = new BoomExportService();
            byte[] bytes = export.Export();
            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(bytes)
            };
            response.Content.Headers.ContentType =
                new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.Content.Headers.ContentDisposition =
                new ContentDispositionHeaderValue("attachment") { FileName = "分部分项清单.xlsx" };
            return ResponseMessage(response);
        }

        /// <summary>C组Reset：清库→重读两个模板重导→重新生成两张导出表（不依赖残留）</summary>
        [HttpPost, Route("api/cgroup/reset")]
        public IHttpActionResult Reset(string key)
        {
            var expectedKey = System.Configuration.ConfigurationManager.AppSettings["ResetKey"] ?? Repositories.Db.ResetKey;
            if (key != expectedKey) return BadRequest("无效的重置密钥");

            // 1. 清空 C 组全部表（含分部分项）
            _repo.ClearCgroup();
            // 2. 从 bin 重新读取两个输入模板并重新导入
            var r = _import.Import(GongfaFile, MainMatFile);
            var boomN = _boomImport.Import(Path.Combine(BinDir, "分部分项清单.json"));
            // 3. 随后重新生成两张导出表，写回 bin（每次运行结果一致、不依赖残留）
            string gongfaOut = Path.Combine(BinDir, "工法主材汇总表.xlsx");
            string boomOut = Path.Combine(BinDir, "分部分项清单.xlsx");
            File.WriteAllBytes(gongfaOut, new CgroupExportService().ExportGongFa());
            File.WriteAllBytes(boomOut, new BoomExportService().Export());

            return Ok(new
            {
                message = "C组已重置、重导并重新生成两张导出表",
                gongfa = r.GongFaCount,
                boom = boomN,
                files = new[] { "工法主材汇总表.xlsx", "分部分项清单.xlsx" }
            });
        }
    }
}
