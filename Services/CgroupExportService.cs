using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using WebApplication1.Repositories;

namespace WebApplication1.Services
{
    /// <summary>
    /// C组工法主材汇总表导出：以"工法清单模板--3.定额计价工法导出.xlsx"为模板，
    /// 打开模板后按 B列类目名 定位 主材/人工/辅材/机械/设备/其它，按 C列资源名逐行填充，
    /// 样式/合并/列宽/行高/打印设置全部沿用模板（只改值，不动样式）。
    /// 数据全部来自 SQLite 持久化结果（算钱服务 + 工法资源），不直接读输入文件。
    /// </summary>
    public class CgroupExportService
    {
        private readonly CgroupCalcService _calc = new CgroupCalcService();

        public byte[] ExportGongFa()
        {
            var binDir = AppDomain.CurrentDomain.BaseDirectory;
            var tplPath = Path.Combine(binDir, "工法清单模板--3.定额计价工法导出.xlsx");
            if (!File.Exists(tplPath))
                throw new FileNotFoundException("未找到工法导出模板：" + tplPath);

            var list = _calc.CalcAll();

            using (var fs = new FileStream(tplPath, FileMode.Open, FileAccess.Read))
            using (IWorkbook wb = new XSSFWorkbook(fs))
            {
                ISheet sheet = wb.GetSheetAt(0);
                foreach (var gf in list)
                    FillOneGongFa(sheet, gf);

                using (var ms = new MemoryStream())
                {
                    wb.Write(ms);
                    return ms.ToArray();
                }
            }
        }

        // 在模板上填充一个工法：按 B列类目、C列名称定位，只改值不动样式
        private void FillOneGongFa(ISheet sheet, GongFaPrice gf)
        {
            // 该工法各类资源（占位行已在取数时排除），全部来自持久化数据
            var main  = GetResources(gf.GongFaId, "主材");
            var labor = GetResources(gf.GongFaId, "人工");
            var aux   = GetResources(gf.GongFaId, "辅材");
            var equip = GetResources(gf.GongFaId, "设备");
            var mech  = GetResources(gf.GongFaId, "机械");

            int mainIdx = 0, laborIdx = 0, auxIdx = 0, equipIdx = 0, mechIdx = 0;
            string currentCat = null;

            // 标题（A1，模板已合并 A1:J1）
            var titleRow = sheet.GetRow(0) ?? sheet.CreateRow(0);
            var titleCell = titleRow.GetCell(0) ?? titleRow.CreateCell(0);
            titleCell.SetCellValue(gf.Name + "（基础工法）");

            for (int r = 1; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;

                string cat = GetCellString(row, 1);  // B列=类目（主材/人工/辅材/机械/设备/其它）
                string name = GetCellString(row, 2); // C列=名称

                // 类目行：只切换当前类目，本身不填数据
                if (cat.Contains("主材")) { currentCat = "主材"; continue; }
                if (cat.Contains("人工")) { currentCat = "人工"; continue; }
                if (cat.Contains("设备")) { currentCat = "设备"; continue; }
                if (cat.Contains("机械")) { currentCat = "机械"; continue; }
                if (cat.Contains("辅材") || cat.Contains("材料")) { currentCat = "辅材"; continue; }
                if (cat.Contains("其它")) { currentCat = "其它"; continue; }

                // 管理费 / 利润 行（C列名称），填 H 列小计（金额保留2位小数）；G列费率标签沿用模板
                if (name.Contains("管理费")) { SetNumber(row, 7, ToDouble(Math.Round(gf.Glf, 2))); continue; }
                if (name.Contains("利润")) { SetNumber(row, 7, ToDouble(Math.Round(gf.Profit, 2))); continue; }

                if (string.IsNullOrEmpty(name) && currentCat != "主材") continue;

                ResRow it = null;
                switch (currentCat)
                {
                    case "主材":
                        if (mainIdx < main.Count) it = main[mainIdx++];
                        if (it == null) continue;
                        // C名称 D规格 E单位 F收入价 G消耗量 H小计 I工法单位 J综合计价
                        SetStr(row, 2, it.Name);
                        SetStr(row, 3, string.IsNullOrEmpty(it.Spec) ? "-" : it.Spec);
                        SetStr(row, 4, it.Unit);
                        SetNumber(row, 5, ToDouble(it.Price));
                        SetNumber(row, 6, ToDouble(it.Dosage));
                        SetNumber(row, 7, ToDouble(Math.Round(it.Dosage * it.Price, 2)));
                        SetStr(row, 8, gf.Unit ?? "");
                        SetNumber(row, 9, ToDouble(Math.Round(gf.TotalPrice, 2))); // 仅主材资源行展示综合计价
                        break;
                    case "人工":
                        if (laborIdx < labor.Count) it = labor[laborIdx++];
                        FillResource(row, it);
                        break;
                    case "辅材":
                        if (auxIdx < aux.Count) it = aux[auxIdx++];
                        FillResource(row, it);
                        break;
                    case "设备":
                        if (equipIdx < equip.Count) it = equip[equipIdx++];
                        FillResource(row, it);
                        break;
                    case "机械":
                        if (mechIdx < mech.Count) it = mech[mechIdx++];
                        FillResource(row, it);
                        break;
                }
            }
        }

        // 资源明细行：C名称 D规格 E单位 F收入价(单价) G消耗量(含量) H小计
        private void FillResource(IRow row, ResRow it)
        {
            if (it == null) return;
            SetStr(row, 2, it.Name);
            SetStr(row, 3, string.IsNullOrEmpty(it.Spec) ? "-" : it.Spec);
            SetStr(row, 4, it.Unit);
            SetNumber(row, 5, ToDouble(it.Price));
            SetNumber(row, 6, ToDouble(it.Dosage));
            SetNumber(row, 7, ToDouble(Math.Round(it.Dosage * it.Price, 2)));
        }

        // 从数据库读取某工法某类资源（占位行 IsPlaceholder=1 不进导出明细，但仍保留在库中）
        private List<ResRow> GetResources(long gongFaId, string type)
        {
            var list = new List<ResRow>();
            using (var conn = new Microsoft.Data.Sqlite.SqliteConnection(Db.ConnectionString))
            {
                conn.Open();
                using (var cmd = new Microsoft.Data.Sqlite.SqliteCommand(
                    "SELECT Name,Spec,Unit,Dosage,Price FROM GongFaResource WHERE GongFaId=$gf AND ResourceType=$t AND IsPlaceholder=0 ORDER BY OrderNo", conn))
                {
                    cmd.Parameters.AddWithValue("$gf", gongFaId);
                    cmd.Parameters.AddWithValue("$t", type);
                    using (var dr = cmd.ExecuteReader())
                        while (dr.Read())
                            list.Add(new ResRow
                            {
                                Name = dr.IsDBNull(0) ? "" : dr.GetString(0),
                                Spec = dr.IsDBNull(1) ? "" : dr.GetString(1),
                                Unit = dr.IsDBNull(2) ? "" : dr.GetString(2),
                                Dosage = dr.IsDBNull(3) ? 0 : dr.GetDecimal(3),
                                Price = dr.IsDBNull(4) ? 0 : dr.GetDecimal(4)
                            });
                }
            }
            return list;
        }

        private static double ToDouble(decimal d) => (double)d;

        private static void SetStr(IRow row, int colIdx, string v)
        {
            var c = row.GetCell(colIdx) ?? row.CreateCell(colIdx);
            c.SetCellValue(v ?? "");
        }
        private static void SetNumber(IRow row, int colIdx, double v)
        {
            var c = row.GetCell(colIdx) ?? row.CreateCell(colIdx);
            c.SetCellValue(v);
        }
        private static string GetCellString(IRow row, int colIdx)
        {
            var c = row.GetCell(colIdx);
            if (c == null) return "";
            if (c.CellType == CellType.String) return c.StringCellValue?.Trim() ?? "";
            if (c.CellType == CellType.Numeric) return c.NumericCellValue.ToString();
            try { return c.ToString()?.Trim() ?? ""; }
            catch { return ""; }
        }

        private class ResRow
        {
            public string Name; public string Spec; public string Unit;
            public decimal Dosage; public decimal Price;
        }
    }
}
