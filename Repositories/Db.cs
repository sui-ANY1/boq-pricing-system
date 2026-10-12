using System;
using Microsoft.Data.Sqlite;

namespace WebApplication1.Repositories
{
    /// <summary>
    /// 数据库基础设施：提供连接，并在首次使用时建表 + 初始化固定工项。
    /// </summary>
    public static class Db
    {
        // 连接字符串：test.db 生成在运行目录(bin)下
        public const string ConnectionString = "Data Source=test.db;";

        /// <summary>固定工项所属的默认项目标识</summary>
        public const string DefaultProjectId = "DEFAULT";

        /// <summary>Reset 接口的内部访问密钥（演示用，生产环境应换成鉴权/内网限制）</summary>
        public const string ResetKey = "reset2026";

        /// <summary>建表 + 初始化固定工项（幂等，可重复调用）</summary>
        public static void Init()
        {
            using (var conn = new SqliteConnection(ConnectionString))
            {
                conn.Open();

                // B组：其他费用
                string createOther = @"
CREATE TABLE IF NOT EXISTS OtherFee (
    Id           INTEGER PRIMARY KEY AUTOINCREMENT,
    ProjectId    TEXT    NOT NULL,
    Name         TEXT    NOT NULL,
    IncomePrice  DECIMAL NULL,
    CostPrice    DECIMAL NOT NULL DEFAULT 0,
    Remark       TEXT    NULL,
    UNIQUE(ProjectId, Name)
);";
                using (var cmd = new SqliteCommand(createOther, conn)) { cmd.ExecuteNonQuery(); }

                // C组：工法组价与分部分项
                string createC = @"
CREATE TABLE IF NOT EXISTS GongFa (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Code TEXT NOT NULL, Name TEXT NOT NULL, Unit TEXT NULL, Part TEXT NULL, UNIQUE(Code)
);
CREATE TABLE IF NOT EXISTS GongFaResource (
    Id INTEGER PRIMARY KEY AUTOINCREMENT, GongFaId INTEGER NOT NULL, OrderNo INTEGER DEFAULT 0,
    ResourceType TEXT NULL, Name TEXT NULL, Spec TEXT NULL, Unit TEXT NULL,
    Dosage DECIMAL NULL, Price DECIMAL NULL, IsPlaceholder INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS MainMaterial (
    Id INTEGER PRIMARY KEY AUTOINCREMENT, Name TEXT NOT NULL, Spec TEXT NULL, Unit TEXT NULL,
    IncomePrice DECIMAL NULL, Dosage DECIMAL NULL, UNIQUE(Name, Spec)
);
CREATE TABLE IF NOT EXISTS FeeRule (
    Id INTEGER PRIMARY KEY AUTOINCREMENT, Name TEXT NOT NULL, Base TEXT NULL, Rate DECIMAL NULL
);
CREATE TABLE IF NOT EXISTS BoomItem (
    Id INTEGER PRIMARY KEY AUTOINCREMENT, Path TEXT NULL, GongFaCode TEXT NULL, Name TEXT NULL,
    MaterialName TEXT NULL, Spec TEXT NULL, Unit TEXT NULL, Quantity DECIMAL NULL,
    IncomeMaterial DECIMAL NULL, IncomeConstruct DECIMAL NULL, IncomeTotal DECIMAL NULL,
    CostMaterial DECIMAL NULL, CostConstruct DECIMAL NULL, CostTotal DECIMAL NULL
);";
                using (var cmdC = new SqliteCommand(createC, conn)) { cmdC.ExecuteNonQuery(); }

                // 初始化固定 6 个工项到默认项目
                string[] fixedItems =
                {
                    "岩土工程", "建筑工程", "技术措施费",
                    "安全文明施工及其他措施项目", "规费", "其它工程"
                };
                foreach (var name in fixedItems)
                {
                    string insertSql = "INSERT OR IGNORE INTO OtherFee(ProjectId, Name, IncomePrice, CostPrice, Remark) VALUES($pid, $name, NULL, 0, NULL);";
                    using (var cmd = new SqliteCommand(insertSql, conn))
                    {
                        cmd.Parameters.AddWithValue("$pid", DefaultProjectId);
                        cmd.Parameters.AddWithValue("$name", name);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        /// <summary>清空数据并重建表（Reset，保证运行不依赖残留数据）</summary>
        public static void Reset()
        {
            using (var conn = new SqliteConnection(ConnectionString))
            {
                conn.Open();
                string dropSql = "DROP TABLE IF EXISTS BoomItem; DROP TABLE IF EXISTS FeeRule; DROP TABLE IF EXISTS MainMaterial; DROP TABLE IF EXISTS GongFaResource; DROP TABLE IF EXISTS GongFa; DROP TABLE IF EXISTS OtherFee;";
                using (var cmd = new SqliteCommand(dropSql, conn)) { cmd.ExecuteNonQuery(); }
            }
            Init();
        }
    }
}
