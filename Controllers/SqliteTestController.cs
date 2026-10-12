using System;
using Microsoft.Data.Sqlite;
using System.Web.Http;

namespace WebApplication1.Controllers
{
    public class SqliteTestController : ApiController
    {
        [HttpGet]
        public IHttpActionResult TestConnect()
        {
            string connStr = "Data Source=test.db;";
            try
            {
                using (SqliteConnection conn = new SqliteConnection(connStr))
                {
                    conn.Open();
                    return Ok("✅ SQLite连接成功！驱动正常");
                }
            }
            catch (Exception ex)
            {
                return InternalServerError(ex);
            }
        }
    }
}
