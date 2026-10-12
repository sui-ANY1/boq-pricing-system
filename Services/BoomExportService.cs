using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using WebApplication1.Models;
using WebApplication1.Repositories;

namespace WebApplication1.Services
{
    /// <summary>
    /// C组分部分项清单导出：以"建筑单体1#分部分项清单"为模板，打开后清掉样例行，
    /// 按数据库持久化数据填充（表头/合并/列宽/边框/打印设置沿用模板）。
    /// 合价=单价×工程量；利润率=(收入-成本)/收入；关联不到工法在备注列给可定位提示。
    /// </summary>
    public class BoomExportService
    {
        private readonly CgroupRepository _repo = new CgroupRepository();

        public byte[] Export()
        {
            var list = _repo.ListBoom();
            var knownCodes = LoadKnownGongFaCodes();

            var binDir = AppDomain.CurrentDomain.BaseDirectory;
            var tplPath = Path.Combine(binDir, "建筑单体1#分部分项清单-模板.xlsx");
            if (!File.Exists(tplPath))
                throw new FileNotFoundException("未找到分部分项清单模板：" + tplPath);

            using (var fs = new FileStream(tplPath, FileMode.Open, FileAccess.Read))
            using (IWorkbook wb = new XSSFWorkbook(fs))
            {
                ISheet sheet = wb.GetSheetAt(0);

                // 1. 动态定位（不写死行号）：
                //    数据起始行=第一个 B列为纯数字编码的行（自动跳过 楼层/分部/功能分区 层级行）；
                //    合计行=A列含"合计"的行。
                int dataStart = -1, totalRowIdx = -1;
                for (int r = 3; r <= sheet.LastRowNum; r++)
                {
                    var row = sheet.GetRow(r);
                    if (row == null) continue;
                    string b = GetString(row, 1);
                    string a = GetString(row, 0);
                    if (dataStart < 0 && IsCode(b)) dataStart = r;
                    if (a.Contains("合计")) { totalRowIdx = r; break; }
                }
                if (dataStart < 0) dataStart = 5;

                // 2. 用数据库数据覆盖 dataStart 起的数据行（样式沿用模板）
                decimal sumIncome = 0, sumCost = 0;
                for (int i = 0; i < list.Count; i++)
                {
                    var it = list[i];
                    int r = dataStart + i;
                    var row = sheet.GetRow(r) ?? sheet.CreateRow(r);

                    decimal incomeUnit = it.IncomeTotal ?? 0;
                    decimal costUnit = it.CostTotal ?? 0;
                    decimal qty = it.Quantity ?? 0;
                    decimal incomeAmount = incomeUnit * qty;
                    decimal costAmount = costUnit * qty;
                    decimal profitRate = incomeUnit != 0 ? (incomeUnit - costUnit) / incomeUnit : 0;
                    sumIncome += incomeAmount;
                    sumCost += costAmount;

                    SetInt(row, 0, i + 1);
                    SetStr(row, 1, it.GongFaCode ?? "");
                    SetStr(row, 2, it.Name ?? "");
                    SetStr(row, 3, it.MaterialName ?? "");
                    SetStr(row, 4, it.Spec ?? "");
                    SetStr(row, 5, it.Unit ?? "");
                    SetNum(row, 6, (double)qty);
                    SetMoney(row, 7, it.IncomeMaterial ?? 0);
                    SetMoney(row, 8, it.IncomeConstruct ?? 0);
                    SetMoney(row, 9, incomeUnit);
                    SetMoney(row, 10, it.CostMaterial ?? 0);
                    SetMoney(row, 11, it.CostConstruct ?? 0);
                    SetMoney(row, 12, costUnit);
                    SetMoney(row, 13, incomeAmount);
                    SetMoney(row, 14, costAmount);
                    SetStr(row, 15, FormatPercent(profitRate)); // P列模板为文本百分比

                    // 关联不到工法：备注列给可定位提示，其余数据照常导出
                    SetStr(row, 16,
                        (!string.IsNullOrEmpty(it.GongFaCode) && !knownCodes.Contains(it.GongFaCode))
                            ? "⚠ 未关联到工法：" + it.GongFaCode
                            : "");
                }

                // 3. 把"合计"行上移到数据末尾，删除中间多余样例行
                int newTotal = dataStart + list.Count;
                if (totalRowIdx > newTotal)
                {
                    sheet.ShiftRows(totalRowIdx, sheet.LastRowNum, -(totalRowIdx - newTotal));
                    totalRowIdx = newTotal;
                }
                else if (totalRowIdx < 0)
                {
                    totalRowIdx = newTotal;
                }

                // 4. 合计行：汇总收入合价/成本合价/利润率；清空其后残留样例行，避免重复
                var tr = sheet.GetRow(totalRowIdx) ?? sheet.CreateRow(totalRowIdx);
                SetStr(tr, 0, "合计");
                SetMoney(tr, 13, sumIncome);
                SetMoney(tr, 14, sumCost);
                SetStr(tr, 15, FormatPercent(sumIncome != 0 ? (sumIncome - sumCost) / sumIncome : 0));
                for (int r = totalRowIdx + 1; r <= sheet.LastRowNum; r++)
                {
                    var rr = sheet.GetRow(r);
                    if (rr == null) continue;
                    for (int c = 0; c <= 16; c++)
                    {
                        var cell = rr.GetCell(c);
                        if (cell != null) cell.SetCellValue("");
                    }
                }

                using (var ms = new MemoryStream())
                {
                    wb.Write(ms);
                    return ms.ToArray();
                }
            }
        }

        // 是否为纯数字项目编码（清单项行）
        private static bool IsCode(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            string t = s.Trim();
            if (t.Length < 6) return false;
            foreach (char ch in t)
                if (!char.IsDigit(ch)) return false;
            return true;
        }

        // 利润率转模板同款文本百分比，如 0.1875 -> "18.75%"
        private static string FormatPercent(decimal rate)
            => (rate * 100m).ToString("0.00") + "%";

        private HashSet<string> LoadKnownGongFaCodes()
        {
            var set = new HashSet<string>();
            using (var conn = new Microsoft.Data.Sqlite.SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                using (var cmd = new Microsoft.Data.Sqlite.SqliteCommand("SELECT Code FROM GongFa", conn))
                using (var dr = cmd.ExecuteReader())
                    while (dr.Read())
                        set.Add(dr.IsDBNull(0) ? "" : dr.GetString(0));
            }
            return set;
        }

        private static string GetString(IRow row, int idx)
        {
            var c = row.GetCell(idx);
            if (c == null) return "";
            try { return c.StringCellValue?.Trim() ?? ""; }
            catch { return c.NumericCellValue.ToString(); }
        }
        private static void SetStr(IRow row, int idx, string v)
        {
            var c = row.GetCell(idx) ?? row.CreateCell(idx);
            c.SetCellValue(v ?? "");
        }
        private static void SetNum(IRow row, int idx, double v)
        {
            var c = row.GetCell(idx) ?? row.CreateCell(idx);
            c.SetCellValue(v);
        }
        // 金额列：保留两位小数
        private static void SetMoney(IRow row, int idx, decimal v)
        {
            var c = row.GetCell(idx) ?? row.CreateCell(idx);
            c.SetCellValue((double)Math.Round(v, 2));
        }
        private static void SetInt(IRow row, int idx, int v)
        {
            var c = row.GetCell(idx) ?? row.CreateCell(idx);
            c.SetCellValue(v);
        }
    }
}
