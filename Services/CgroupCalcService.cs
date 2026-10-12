using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using Microsoft.Data.Sqlite;
using WebApplication1.Models;

namespace WebApplication1.Services
{
    /// <summary>单个工法的组价结果（现场算出来的）</summary>
    public class GongFaPrice
    {
        public long GongFaId { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public string Unit { get; set; }
        public decimal Labor { get; set; }          // 人工费
        public decimal AuxMaterial { get; set; }    // 辅材费
        public decimal Equipment { get; set; }      // 设备费
        public decimal Mechanical { get; set; }     // 机械费
        public decimal MainMaterialFee { get; set; } // 主材费
        public decimal MaterialFee { get; set; }    // 材料费=主材+辅材+设备（A组口径）
        public decimal DirectFee { get; set; }      // 直接费=人工+材料+机械
        public decimal Glf { get; set; }            // 管理费
        public decimal Profit { get; set; }         // 利润
        public decimal TotalPrice { get; set; }     // 综合单价=直接费+管理费+利润
    }

    /// <summary>
    /// C组算钱：资源金额=含量×有效单价；按 人工/主材/辅材/设备/机械 分类汇总（A组口径）；
    /// 管理费/利润按库里配置的计价基数表达式×费率逐条求值（基数或费率一改，结果随之变）。
    /// 占位资源(无单价)只保留业务信息、不计金额。主材费取本工法挂接的主材资源，可追溯、非全表样例。
    /// </summary>
    public class CgroupCalcService
    {
        /// <summary>计算所有工法的组价</summary>
        public List<GongFaPrice> CalcAll()
        {
            var list = new List<GongFaPrice>();
            using (var conn = new SqliteConnection(DbConnection()))
            {
                conn.Open();
                var gfs = new List<GongFa>();
                using (var cmd = new SqliteCommand("SELECT Id,Code,Name,Unit,Part FROM GongFa", conn))
                using (var dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                        gfs.Add(new GongFa
                        {
                            Id = dr.GetInt64(0),
                            Code = dr.GetString(1),
                            Name = dr.GetString(2),
                            Unit = dr.IsDBNull(3) ? "" : dr.GetString(3),
                            Part = dr.IsDBNull(4) ? "" : dr.GetString(4)
                        });
                }
                foreach (var gf in gfs)
                    list.Add(CalcOne(conn, gf));
            }
            return list;
        }

        private static string DbConnection() => Repositories.Db.ConnectionString;

        private GongFaPrice CalcOne(SqliteConnection conn, GongFa gf)
        {
            var p = new GongFaPrice { GongFaId = gf.Id, Code = gf.Code, Name = gf.Name, Unit = gf.Unit };

            // 1. 读本工法资源明细，按类型求和（Price空=占位，跳过不计）
            using (var cmd = new SqliteCommand(
                "SELECT ResourceType,Dosage,Price FROM GongFaResource WHERE GongFaId=$id ORDER BY OrderNo", conn))
            {
                cmd.Parameters.AddWithValue("$id", gf.Id);
                using (var dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        if (dr.IsDBNull(2)) continue; // 无单价=占位/缺价，不计金额
                        string type = dr.IsDBNull(0) ? "" : dr.GetString(0);
                        decimal dosage = dr.IsDBNull(1) ? 0 : dr.GetDecimal(1);
                        decimal price = dr.GetDecimal(2);
                        decimal amount = dosage * price;

                        if (type == "人工") p.Labor += amount;
                        else if (type == "机械") p.Mechanical += amount;
                        else if (type == "主材") p.MainMaterialFee += amount;
                        else if (type == "设备") p.Equipment += amount;
                        else p.AuxMaterial += amount; // 辅材（默认材料桶）
                    }
                }
            }

            // 2. 材料费（含主材、辅材、设备）与直接费（A组口径）
            p.MaterialFee = p.MainMaterialFee + p.AuxMaterial + p.Equipment;
            p.DirectFee = p.Labor + p.MaterialFee + p.Mechanical;

            // 3. 取费规则按 Id 顺序（先管理费后利润），把中文基数token替换成数字后求值
            using (var cmd = new SqliteCommand("SELECT Name,Base,Rate FROM FeeRule ORDER BY Id", conn))
            using (var dr = cmd.ExecuteReader())
            {
                while (dr.Read())
                {
                    string name = dr.GetString(0);
                    string expr = dr.IsDBNull(1) ? "" : dr.GetString(1);
                    decimal rate = dr.IsDBNull(2) ? 0 : dr.GetDecimal(2);

                    // 用固定小数点格式，避免区域设置把小数点变成逗号导致算式错误
                    string replaced = expr
                        .Replace("分部分项人工费", p.Labor.ToString(CultureInfo.InvariantCulture))
                        .Replace("分部分项材料费", p.MaterialFee.ToString(CultureInfo.InvariantCulture))
                        .Replace("分部分项机械费", p.Mechanical.ToString(CultureInfo.InvariantCulture))
                        .Replace("分部分项主材费", p.MainMaterialFee.ToString(CultureInfo.InvariantCulture))
                        .Replace("分部分项设备费", p.Equipment.ToString(CultureInfo.InvariantCulture))
                        .Replace("分部分项直接费", p.DirectFee.ToString(CultureInfo.InvariantCulture))
                        .Replace("分部分项管理费", p.Glf.ToString(CultureInfo.InvariantCulture));

                    decimal fee = Eval(replaced) * rate;

                    if (name.Contains("管理")) p.Glf = fee;
                    else if (name.Contains("利润")) p.Profit = fee;
                }
            }

            // 4. 综合单价 = 直接费 + 管理费 + 利润
            p.TotalPrice = p.DirectFee + p.Glf + p.Profit;
            return p;
        }

        /// <summary>用 DataTable.Compute 算四则运算表达式；无法求值按 0 处理（不抛异常）</summary>
        private decimal Eval(string expr)
        {
            if (string.IsNullOrWhiteSpace(expr)) return 0;
            try
            {
                using (var t = new DataTable())
                {
                    object v = t.Compute(expr, "");
                    return Convert.ToDecimal(v);
                }
            }
            catch { return 0; }
        }
    }
}
