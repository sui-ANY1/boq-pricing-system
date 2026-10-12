using System.Collections.Generic;
using Microsoft.Data.Sqlite;
using WebApplication1.Models;

namespace WebApplication1.Repositories
{
    /// <summary>
    /// 其他费用 仓储层：只负责和 SQLite 打交道（增删改查 SQL）。
    /// </summary>
    public class OtherFeeRepository
    {
        private OtherFee Map(SqliteDataReader r)
        {
            return new OtherFee
            {
                Id = (long)r["Id"],
                ProjectId = r["ProjectId"].ToString(),
                Name = r["Name"].ToString(),
                IncomePrice = r["IncomePrice"] == System.DBNull.Value ? (decimal?)null : System.Convert.ToDecimal(r["IncomePrice"]),
                CostPrice = System.Convert.ToDecimal(r["CostPrice"]),
                Remark = r["Remark"] == System.DBNull.Value ? null : r["Remark"].ToString()
            };
        }

        /// <summary>新增，返回自增Id</summary>
        public long Insert(OtherFee fee)
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                string sql = @"INSERT INTO OtherFee(ProjectId, Name, IncomePrice, CostPrice, Remark)
                               VALUES($pid, $name, $income, $cost, $remark);
                               SELECT last_insert_rowid();";
                using (var cmd = new SqliteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("$pid", fee.ProjectId);
                    cmd.Parameters.AddWithValue("$name", fee.Name);
                    cmd.Parameters.AddWithValue("$income", (object)fee.IncomePrice ?? System.DBNull.Value);
                    cmd.Parameters.AddWithValue("$cost", fee.CostPrice);
                    cmd.Parameters.AddWithValue("$remark", (object)fee.Remark ?? System.DBNull.Value);
                    return (long)cmd.ExecuteScalar();
                }
            }
        }

        /// <summary>按项目查全部</summary>
        public List<OtherFee> ListByProject(string projectId)
        {
            var list = new List<OtherFee>();
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                string sql = "SELECT * FROM OtherFee WHERE ProjectId = $pid ORDER BY Id";
                using (var cmd = new SqliteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("$pid", projectId);
                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read()) list.Add(Map(r));
                    }
                }
            }
            return list;
        }

        public OtherFee GetById(long id)
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                string sql = "SELECT * FROM OtherFee WHERE Id = $id";
                using (var cmd = new SqliteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("$id", id);
                    using (var r = cmd.ExecuteReader())
                    {
                        if (r.Read()) return Map(r);
                        return null;
                    }
                }
            }
        }

        /// <summary>判断同项目下名称是否重复（excludeId 用于修改时排除自己）</summary>
        public bool ExistsByName(string projectId, string name, long excludeId = 0)
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                string sql = @"SELECT COUNT(1) FROM OtherFee
                               WHERE ProjectId = $pid AND Name = $name AND Id <> $exclude";
                using (var cmd = new SqliteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("$pid", projectId);
                    cmd.Parameters.AddWithValue("$name", name);
                    cmd.Parameters.AddWithValue("$exclude", excludeId);
                    return System.Convert.ToInt64(cmd.ExecuteScalar()) > 0;
                }
            }
        }

        /// <summary>按项目+名称精确查单条（导入时复用，避免存在性查询和取记录各开一次连接）</summary>
        public OtherFee FindByName(string projectId, string name)
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                string sql = "SELECT * FROM OtherFee WHERE ProjectId = $pid AND Name = $name LIMIT 1";
                using (var cmd = new SqliteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("$pid", projectId);
                    cmd.Parameters.AddWithValue("$name", name);
                    using (var r = cmd.ExecuteReader())
                    {
                        if (r.Read()) return Map(r);
                        return null;
                    }
                }
            }
        }

        public void Update(OtherFee fee)
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                string sql = @"UPDATE OtherFee SET Name=$name, IncomePrice=$income, CostPrice=$cost, Remark=$remark
                               WHERE Id=$id";
                using (var cmd = new SqliteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("$name", fee.Name);
                    cmd.Parameters.AddWithValue("$income", (object)fee.IncomePrice ?? System.DBNull.Value);
                    cmd.Parameters.AddWithValue("$cost", fee.CostPrice);
                    cmd.Parameters.AddWithValue("$remark", (object)fee.Remark ?? System.DBNull.Value);
                    cmd.Parameters.AddWithValue("$id", fee.Id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void Delete(long id)
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqliteCommand("DELETE FROM OtherFee WHERE Id=$id", conn))
                {
                    cmd.Parameters.AddWithValue("$id", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}
