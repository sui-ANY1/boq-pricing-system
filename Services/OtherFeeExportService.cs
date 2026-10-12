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
    /// 导出"其他费用清单"：打开样例模板"导出其他费用清单.xlsx"填数据，
    /// 保留模板的标题/表头/合并/列宽/边框/打印设置（符合要求4.5）。
    /// 金额、利润、税金、利润率全部从数据库现算，可追溯到持久化数据。
    /// </summary>
    public class OtherFeeExportService
    {
        private readonly OtherFeeRepository _repo = new OtherFeeRepository();

        // 列（0基）：A序号 B名称 C收入 D成本 E利润 F建面 G单方收入 H单方成本 I利润率 J备注
        private const int COL_SEQ = 0, COL_NAME = 1, COL_INC = 2, COL_COST = 3, COL_PROFIT = 4,
                          COL_AREA = 5, COL_INC_UNIT = 6, COL_COST_UNIT = 7, COL_RATE = 8, COL_REMARK = 9;

        public byte[] Export(string projectId, string projectName)
        {
            var fixedFees = _repo.ListByProject(Db.DefaultProjectId);
            var customFees = string.Equals(projectId, Db.DefaultProjectId, StringComparison.OrdinalIgnoreCase)
                ? new List<OtherFee>()
                : _repo.ListByProject(projectId);

            var binDir = AppDomain.CurrentDomain.BaseDirectory;
            var tplPath = Path.Combine(binDir, "导出其他费用清单.xlsx");
            if (!File.Exists(tplPath))
                throw new FileNotFoundException("导出模板文件缺失：导出其他费用清单.xlsx", tplPath);

            using (var fs = new FileStream(tplPath, FileMode.Open, FileAccess.Read))
            using (IWorkbook workbook = new XSSFWorkbook(fs))
            {
                ISheet sheet = workbook.GetSheetAt(0);

                // 标题
                SetStr(sheet, 0, 0, projectName + " - 项目汇总清单");

                // 1) 6个固定工项：模板第4~9行（idx3~8）
                decimal sectionInc = 0, sectionCost = 0;
                for (int i = 0; i < fixedFees.Count; i++)
                {
                    var f = fixedFees[i];
                    decimal inc = f.IncomePrice ?? 0m, cost = f.CostPrice;
                    WriteFeeRow(sheet, 3 + i, (i + 1).ToString(), f.Name, inc, cost, f.Remark);
                    sectionInc += inc; sectionCost += cost;
                }

                // 2) 其他费用：从模板第10行（idx9）起。模板原只留1行样例，按实际N行调整
                int n = customFees.Count;
                int firstDataRow = 9; // 第10行
                // 二/三/四/五 当前在 idx10..13；若N>1，把它们往下挪 (N-1) 行腾出其他费用行
                if (n > 1)
                    sheet.ShiftRows(10, sheet.LastRowNum, n - 1);
                else if (n <= 0)
                    sheet.ShiftRows(10, sheet.LastRowNum, -1); // 删掉那行样例其他费用

                for (int i = 0; i < n; i++)
                {
                    var f = customFees[i];
                    decimal inc = f.IncomePrice ?? 0m, cost = f.CostPrice;
                    WriteFeeRow(sheet, firstDataRow + i, (fixedFees.Count + i + 1).ToString(),
                        f.Name, inc, cost, f.Remark);
                    sectionInc += inc; sectionCost += cost;
                }

                // 3) 按B列名称定位 一/二/三/四/五 行，回填汇总
                decimal designInc = 0, designCost = 0;
                decimal noTaxInc = sectionInc + designInc, noTaxCost = sectionCost + designCost;
                decimal taxInc = noTaxInc * 0.09m, taxCost = noTaxCost * 0.09m;
                decimal totalInc = noTaxInc + taxInc, totalCost = noTaxCost + taxCost;

                for (int r = 0; r <= sheet.LastRowNum; r++)
                {
                    var row = sheet.GetRow(r);
                    if (row == null) continue;
                    string b = GetStr(row, COL_NAME);
                    if (b.Contains("建设安装工程费")) WriteSection(row, sectionInc, sectionCost);
                    else if (b.Contains("设计费")) WriteSection(row, designInc, designCost);
                    else if (b.Contains("不含税合计")) WriteSection(row, noTaxInc, noTaxCost);
                    else if (b.Contains("税金")) WriteTax(row, taxInc, taxCost);
                    else if (b.Contains("含税合计")) WriteSection(row, totalInc, totalCost);
                }

                // 4) 打印设置：A4 横向、适配一页宽、页边距
                sheet.PrintSetup.PaperSize = (short)PaperSize.A4;
                sheet.PrintSetup.Landscape = true;
                sheet.FitToPage = true;
                sheet.PrintSetup.FitWidth = 1;
                sheet.PrintSetup.FitHeight = 0;
                sheet.SetMargin(MarginType.LeftMargin, 0.7);
                sheet.SetMargin(MarginType.RightMargin, 0.7);
                sheet.SetMargin(MarginType.TopMargin, 0.75);
                sheet.SetMargin(MarginType.BottomMargin, 0.75);

                using (var ms = new MemoryStream())
                {
                    workbook.Write(ms);
                    return ms.ToArray();
                }
            }
        }

        // 写一条费用行（样式沿用模板，只改值）
        private void WriteFeeRow(ISheet sheet, int r, string seq, string name,
            decimal inc, decimal cost, string remark)
        {
            var row = sheet.GetRow(r) ?? sheet.CreateRow(r);
            SetStr(row, COL_SEQ, seq);
            SetStr(row, COL_NAME, name);
            SetNum(row, COL_INC, inc);
            SetNum(row, COL_COST, cost);
            SetNum(row, COL_PROFIT, inc - cost);
            SetNum(row, COL_AREA, 0m);
            SetNum(row, COL_INC_UNIT, 0m);
            SetNum(row, COL_COST_UNIT, 0m);
            SetRate(row, COL_RATE, inc - cost, inc);
            SetStr(row, COL_REMARK, remark ?? "");
        }

        // 汇总行：C收入 D成本 E利润 I利润率（按汇总重算）
        private void WriteSection(IRow row, decimal inc, decimal cost)
        {
            SetNum(row, COL_INC, inc);
            SetNum(row, COL_COST, cost);
            SetNum(row, COL_PROFIT, inc - cost);
            SetRate(row, COL_RATE, inc - cost, inc);
        }

        // 税金行：C收入税 D成本税 E税差（F~I 留空，与样例一致）
        private void WriteTax(IRow row, decimal inc, decimal cost)
        {
            SetNum(row, COL_INC, inc);
            SetNum(row, COL_COST, cost);
            SetNum(row, COL_PROFIT, inc - cost);
        }

        private static void SetStr(ISheet sheet, int r, int c, string v)
        {
            var row = sheet.GetRow(r) ?? sheet.CreateRow(r);
            SetStr(row, c, v);
        }
        private static void SetStr(IRow row, int c, string v)
        {
            var cell = row.GetCell(c) ?? row.CreateCell(c);
            cell.SetCellValue(v ?? "");
        }
        private static void SetNum(IRow row, int c, decimal v)
        {
            var cell = row.GetCell(c) ?? row.CreateCell(c);
            cell.SetCellValue(Convert.ToDouble(Math.Round(v, 2)));
        }
        private static void SetRate(IRow row, int c, decimal profit, decimal income)
        {
            var cell = row.GetCell(c) ?? row.CreateCell(c);
            cell.SetCellValue(income == 0 ? 0 : Convert.ToDouble(profit / income));
        }
        private static string GetStr(IRow row, int c)
        {
            var cell = row.GetCell(c);
            if (cell == null) return "";
            switch (cell.CellType)
            {
                case CellType.String: return cell.StringCellValue?.Trim() ?? "";
                case CellType.Numeric: return cell.NumericCellValue.ToString();
                case CellType.Boolean: return cell.BooleanCellValue.ToString();
                case CellType.Formula:
                    return cell.CachedFormulaResultType == CellType.String
                        ? (cell.StringCellValue?.Trim() ?? "")
                        : cell.NumericCellValue.ToString();
                default: return "";
            }
        }
    }
}
