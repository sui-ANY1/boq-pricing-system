using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using WebApplication1.Models;
using WebApplication1.Repositories;

namespace WebApplication1.Services
{
    public class ImportFailure
    {
        public int Row { get; set; }
        public string Reason { get; set; }
    }

    public class ImportReport
    {
        public int InsertCount { get; set; }
        public int OverwriteCount { get; set; }
        public List<ImportFailure> Failures { get; set; } = new List<ImportFailure>();
    }

    /// <summary>
    /// NPOI 解析“其他费用清单”Excel 并入库。
    /// 规则（对应需求 2.6/2.7）：
    ///  - 按表头“名称关键字”定位列，不写死列索引；同时兼容简版（费用名称/描述）与富格式（工项名称/序号/备注）。
    ///  - 只识别“一 建设安装工程费”章节内、序号为数字、且名称不属于 6 个固定工项的行；
    ///    标题行、表头行、章节行（一~五）、固定工项行、合计/税金行一律静默跳过（不计失败）。
    ///  - 兼容文本/数值/公式/空；收入空=null，成本空=无效行；脏数据不中断，记行号+原因；同名覆盖更新。
    /// 列映射为方法内局部变量，服务无状态、可安全并发/单例。
    /// </summary>
    public class OtherFeeImportService
    {
        private readonly OtherFeeRepository _repo = new OtherFeeRepository();

        // 2.6：6个固定工项，不属于“其他费用”，导入时跳过（不是失败行）
        private static readonly HashSet<string> FixedItems = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "岩土工程", "建筑工程", "技术措施费",
            "安全文明施工及其他措施项目", "规费", "其它工程"
        };

        public ImportReport Import(Stream fileStream, string defaultProjectId)
        {
            var report = new ImportReport();

            using (IWorkbook workbook = new XSSFWorkbook(fileStream))
            {
                ISheet sheet = workbook.GetSheetAt(0);

                // 1. 找表头行：前10行内找含“费用名称”或“工项名称”的行
                IRow headerRow = null;
                int headerRowIdx = -1;
                for (int i = 0; i <= sheet.LastRowNum && i < 10; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null) continue;
                    foreach (var cellObj in row.Cells)
                    {
                        var t = GetString((ICell)cellObj);
                        if (t.Contains("费用名称") || t.Contains("工项名称"))
                        {
                            headerRow = row;
                            headerRowIdx = i;
                            break;
                        }
                    }
                    if (headerRow != null) break;
                }
                if (headerRow == null)
                    throw new Exception("未找到表头行（缺少“费用名称/工项名称”列）");

                // 2. 按表头关键字建列映射（不写死列索引，列顺序变化也能识别）
                int seqCol = -1, nameCol = -1, pidCol = -1, incCol = -1, costCol = -1, remarkCol = -1;
                foreach (var cellObj in headerRow.Cells)
                {
                    var cell = (ICell)cellObj;
                    string t = GetString(cell);
                    if (t.Length == 0) continue;
                    int idx = cell.ColumnIndex;
                    if (t.Contains("序号")) seqCol = idx;
                    else if (t.Contains("费用名称") || t.Contains("工项名称")) nameCol = idx;
                    else if (t.Contains("项目Id") || t.Contains("项目编号") || t.Contains("项目ID")) pidCol = idx;
                    else if (t.Contains("收入价")) incCol = idx;
                    else if (t.Contains("成本价")) costCol = idx;
                    else if (t.Contains("备注") || t.Trim() == "描述") remarkCol = idx;
                }
                if (nameCol < 0)
                    throw new Exception("表头缺少“费用名称/工项名称”列");

                // 3. 逐行解析数据
                for (int r = headerRowIdx + 1; r <= sheet.LastRowNum; r++)
                {
                    var row = sheet.GetRow(r);
                    if (row == null || IsRowEmpty(row)) continue;
                    int rowNo = r + 1; // Excel 行号从1开始

                    string name = GetString(row.GetCell(nameCol));

                    // 2.6：固定工项行不是其他费用，跳过（不计失败）
                    if (FixedItems.Contains(name.Trim())) continue;

                    // 2.6：富格式下序号为中文（一/二/三/四/五）的是章节/合计行，跳过
                    if (seqCol >= 0)
                    {
                        string seq = GetString(row.GetCell(seqCol));
                        if (seq.Length > 0 && !decimal.TryParse(seq, out _)) continue;
                    }

                    // 2.6：章节/合计/税金行（按名称兜底识别），跳过
                    if (IsSectionOrTotal(name)) continue;

                    string projectId = pidCol >= 0 ? GetString(row.GetCell(pidCol)) : "";
                    if (string.IsNullOrWhiteSpace(projectId)) projectId = defaultProjectId;

                    decimal? income = incCol >= 0 ? GetDecimal(row.GetCell(incCol)) : null;
                    decimal? cost = costCol >= 0 ? GetDecimal(row.GetCell(costCol)) : null;
                    string remark = remarkCol >= 0 ? GetString(row.GetCell(remarkCol)) : "";

                    // ---- 业务校验（仅对真正的“其他费用”数据行）----
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        report.Failures.Add(new ImportFailure { Row = rowNo, Reason = "费用名称为空" });
                        continue;
                    }
                    if (string.IsNullOrWhiteSpace(projectId))
                    {
                        report.Failures.Add(new ImportFailure { Row = rowNo, Reason = "项目Id为空" });
                        continue;
                    }
                    if (cost == null)
                    {
                        report.Failures.Add(new ImportFailure { Row = rowNo, Reason = "成本价为空或非数字" });
                        continue;
                    }
                    if (cost < 0)
                    {
                        report.Failures.Add(new ImportFailure { Row = rowNo, Reason = "成本价为负数" });
                        continue;
                    }
                    if (income.HasValue && income.Value < 0)
                    {
                        report.Failures.Add(new ImportFailure { Row = rowNo, Reason = "收入价为负数" });
                        continue;
                    }

                    var fee = new OtherFee
                    {
                        ProjectId = projectId,
                        Name = name,
                        IncomePrice = income, // 空收入保留为null
                        CostPrice = cost.Value,
                        Remark = remark
                    };

                    // ---- 同名覆盖：一次 SQL 取记录，与数据库 Name = $name 一致，精确匹配 ----
                    var old = _repo.FindByName(projectId, name);
                    if (old != null)
                    {
                        fee.Id = old.Id;
                        _repo.Update(fee);
                        report.OverwriteCount++;
                    }
                    else
                    {
                        _repo.Insert(fee);
                        report.InsertCount++;
                    }
                }
            } // workbook 在此释放

            return report;
        }

        // 章节/合计/税金行名称识别（2.6 要求跳过，不计失败）
        private static bool IsSectionOrTotal(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            return name.Contains("建设安装工程费")
                || name.Contains("设计费")
                || name.Contains("合计")
                || name.Contains("税金");
        }

        private bool IsRowEmpty(IRow row)
        {
            foreach (var cellObj in row.Cells)
            {
                if (!string.IsNullOrEmpty(GetString((ICell)cellObj))) return false;
            }
            return true;
        }

        // 兼容 文本/数字/公式/空
        private string GetString(ICell cell)
        {
            if (cell == null) return "";
            switch (cell.CellType)
            {
                case CellType.String: return cell.StringCellValue.Trim();
                case CellType.Numeric: return cell.NumericCellValue.ToString();
                case CellType.Boolean: return cell.BooleanCellValue.ToString();
                case CellType.Formula:
                    try { return cell.NumericCellValue.ToString(); }
                    catch { return (cell.StringCellValue ?? "").Trim(); }
                default: return "";
            }
        }

        private decimal? GetDecimal(ICell cell)
        {
            if (cell == null) return null;
            switch (cell.CellType)
            {
                case CellType.Numeric: return System.Convert.ToDecimal(cell.NumericCellValue);
                case CellType.String:
                    string s = cell.StringCellValue.Trim();
                    if (s == "") return null;
                    decimal d;
                    return decimal.TryParse(s, out d) ? (decimal?)d : null;
                case CellType.Formula:
                    try { return System.Convert.ToDecimal(cell.NumericCellValue); }
                    catch { return null; }
                default: return null;
            }
        }
    }
}
