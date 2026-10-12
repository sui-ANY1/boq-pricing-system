using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using WebApplication1.Models;

namespace WebApplication1.Repositories
{
    /// <summary>C组数据访问：工法/资源/主材/取费规则</summary>
    public class CgroupRepository
    {
        /// <summary>清空C组四张表（幂等导入前调用）</summary>
        public void ClearCgroup()
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqliteCommand(
                    "DELETE FROM GongFaResource; DELETE FROM GongFa; DELETE FROM MainMaterial; DELETE FROM FeeRule; DELETE FROM BoomItem;", conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>只清空分部分项</summary>
        public void ClearBoom()
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqliteCommand("DELETE FROM BoomItem;", conn))
                    cmd.ExecuteNonQuery();
            }
        }

        /// <summary>插入分部分项清单项</summary>
        public void InsertBoomItem(BoomItem it)
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqliteCommand(
                    @"INSERT INTO BoomItem(Path,GongFaCode,Name,MaterialName,Spec,Unit,Quantity,
                      IncomeMaterial,IncomeConstruct,IncomeTotal,CostMaterial,CostConstruct,CostTotal)
                      VALUES($p,$code,$n,$mn,$sp,$u,$q,$im,$ic,$it,$cm,$cc,$ct);", conn))
                {
                    cmd.Parameters.AddWithValue("$p", (object)it.Path ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$code", it.GongFaCode ?? "");
                    cmd.Parameters.AddWithValue("$n", it.Name ?? "");
                    cmd.Parameters.AddWithValue("$mn", (object)it.MaterialName ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$sp", (object)it.Spec ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$u", (object)it.Unit ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$q", (object)it.Quantity ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$im", (object)it.IncomeMaterial ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$ic", (object)it.IncomeConstruct ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$it", (object)it.IncomeTotal ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$cm", (object)it.CostMaterial ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$cc", (object)it.CostConstruct ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$ct", (object)it.CostTotal ?? DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>查全部分部分项</summary>
        public List<BoomItem> ListBoom()
        {
            var list = new List<BoomItem>();
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqliteCommand("SELECT GongFaCode,Name,MaterialName,Spec,Unit,Quantity,IncomeMaterial,IncomeConstruct,IncomeTotal,CostMaterial,CostConstruct,CostTotal FROM BoomItem ORDER BY Id", conn))
                using (var dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        list.Add(new BoomItem
                        {
                            GongFaCode = dr.IsDBNull(0) ? "" : dr.GetString(0),
                            Name = dr.IsDBNull(1) ? "" : dr.GetString(1),
                            MaterialName = dr.IsDBNull(2) ? "" : dr.GetString(2),
                            Spec = dr.IsDBNull(3) ? "" : dr.GetString(3),
                            Unit = dr.IsDBNull(4) ? "" : dr.GetString(4),
                            Quantity = dr.IsDBNull(5) ? 0 : dr.GetDecimal(5),
                            IncomeMaterial = dr.IsDBNull(6) ? 0 : dr.GetDecimal(6),
                            IncomeConstruct = dr.IsDBNull(7) ? 0 : dr.GetDecimal(7),
                            IncomeTotal = dr.IsDBNull(8) ? 0 : dr.GetDecimal(8),
                            CostMaterial = dr.IsDBNull(9) ? 0 : dr.GetDecimal(9),
                            CostConstruct = dr.IsDBNull(10) ? 0 : dr.GetDecimal(10),
                            CostTotal = dr.IsDBNull(11) ? 0 : dr.GetDecimal(11)
                        });
                    }
                }
            }
            return list;
        }

        /// <summary>插入工法主表，返回新Id</summary>
        public long InsertGongFa(GongFa gf)
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqliteCommand(
                    "INSERT INTO GongFa(Code, Name, Unit, Part) VALUES($c,$n,$u,$p); SELECT last_insert_rowid();", conn))
                {
                    cmd.Parameters.AddWithValue("$c", gf.Code ?? "");
                    cmd.Parameters.AddWithValue("$n", gf.Name ?? "");
                    cmd.Parameters.AddWithValue("$u", (object)gf.Unit ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$p", (object)gf.Part ?? DBNull.Value);
                    return (long)cmd.ExecuteScalar();
                }
            }
        }

        public void InsertResource(GongFaResource r)
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqliteCommand(
                    @"INSERT INTO GongFaResource(GongFaId, OrderNo, ResourceType, Name, Spec, Unit, Dosage, Price, IsPlaceholder)
                      VALUES($gf,$o,$t,$n,$s,$u,$d,$p,$ph);", conn))
                {
                    cmd.Parameters.AddWithValue("$gf", r.GongFaId);
                    cmd.Parameters.AddWithValue("$o", r.OrderNo);
                    cmd.Parameters.AddWithValue("$t", r.ResourceType ?? "");
                    cmd.Parameters.AddWithValue("$n", r.Name ?? "");
                    cmd.Parameters.AddWithValue("$s", (object)r.Spec ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$u", (object)r.Unit ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$d", (object)r.Dosage ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$p", (object)r.Price ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$ph", r.IsPlaceholder);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>主材按名称+规格 upsert（带价格的来源用）</summary>
        public void UpsertMainMaterial(MainMaterial m)
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqliteCommand(
                    @"INSERT OR REPLACE INTO MainMaterial(Name, Spec, Unit, IncomePrice, Dosage)
                      VALUES($n,$s,$u,$ip,$d);", conn))
                {
                    cmd.Parameters.AddWithValue("$n", m.Name ?? "");
                    cmd.Parameters.AddWithValue("$s", (object)m.Spec ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$u", (object)m.Unit ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$ip", (object)m.IncomePrice ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$d", (object)m.Dosage ?? DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>主材按名称+规格不存在才插入（目录补全用，不覆盖价格）。返回是否新插入。</summary>
        public bool InsertIfNotExistsMainMaterial(MainMaterial m)
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                using (var check = new SqliteCommand("SELECT COUNT(1) FROM MainMaterial WHERE Name=$n AND IFNULL(Spec,'')=IFNULL($s,'');", conn))
                {
                    check.Parameters.AddWithValue("$n", m.Name ?? "");
                    check.Parameters.AddWithValue("$s", (object)m.Spec ?? DBNull.Value);
                    long cnt = (long)check.ExecuteScalar();
                    if (cnt > 0) return false;
                }
                using (var ins = new SqliteCommand(
                    @"INSERT INTO MainMaterial(Name, Spec, Unit, IncomePrice, Dosage)
                      VALUES($n,$s,$u,NULL,NULL);", conn))
                {
                    ins.Parameters.AddWithValue("$n", m.Name ?? "");
                    ins.Parameters.AddWithValue("$s", (object)m.Spec ?? DBNull.Value);
                    ins.Parameters.AddWithValue("$u", (object)m.Unit ?? DBNull.Value);
                    ins.ExecuteNonQuery();
                    return true;
                }
            }
        }

        public void InsertFeeRule(FeeRule f)
        {
            using (var conn = new SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqliteCommand(
                    "INSERT INTO FeeRule(Name, Base, Rate) VALUES($n,$b,$r);", conn))
                {
                    cmd.Parameters.AddWithValue("$n", f.Name ?? "");
                    cmd.Parameters.AddWithValue("$b", (object)f.Base ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$r", (object)f.Rate ?? DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}
