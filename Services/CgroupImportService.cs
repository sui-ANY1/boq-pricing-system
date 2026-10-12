using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using WebApplication1.Models;
using WebApplication1.Repositories;

namespace WebApplication1.Services
{
    /// <summary>
    /// C组导入：解析两个输入Excel，把工法/资源/主材/取费规则写入SQLite。
    /// 工法导入模板：遇到 A列="项目编号" 视为一个新工法块；块内按 序号 区分 人工/材料/机械 三组，
    /// 后续续行沿用"当前分组"。材料行=辅材；价格为空的行标记为占位（保留业务信息但不计金额）。
    /// 主材导入模板：主材既是主档(MainMaterial)，也作为 ResourceType="主材" 的资源挂到对应工法，
    /// 从而让 工法↔主材 的关联在库里可追溯（算钱/导出按工法取，不使用全表第一条这种样例特例）。
    /// </summary>
    public class CgroupImportService
    {
        private readonly CgroupRepository _repo = new CgroupRepository();

        public class CgroupImportResult
        {
            public int GongFaCount { get; set; }
            public int ResourceCount { get; set; }
            public int MainMaterialCount { get; set; }
            public int FeeRuleCount { get; set; }
            public List<long> GongFaIds { get; set; } = new List<long>();
        }

        public CgroupImportResult Import(string gongfaPath, string mainMatPath)
        {
            var result = new CgroupImportResult();
            _repo.ClearCgroup();

            using (var fs = new FileStream(gongfaPath, FileMode.Open, FileAccess.Read))
            using (IWorkbook wb = new XSSFWorkbook(fs))
            {
                var gfSheet = wb.GetSheet("工法导入模板");
                ImportGongFa(gfSheet, result);

                var matSheet = wb.GetSheet("主材导入模板");
                if (matSheet != null)
                    result.MainMaterialCount = ImportMainMaterialWithPrice(matSheet, result.GongFaIds);

                var feeSheet = wb.GetSheet("管理费、利润计算公式");
                if (feeSheet != null)
                    result.FeeRuleCount = ImportFeeRules(feeSheet);
            }

            // 模板1主材目录（只有名称规格，无价格）：仅补充缺失主档，不覆盖已有价格
            if (File.Exists(mainMatPath))
            {
                using (var fs = new FileStream(mainMatPath, FileMode.Open, FileAccess.Read))
                using (IWorkbook wb = new XSSFWorkbook(fs))
                {
                    var ws = wb.GetSheetAt(0);
                    result.MainMaterialCount += ImportMainMaterialCatalog(ws);
                }
            }

            return result;
        }

        // ===== 解析工法块（多个连排），记住当前分组 =====
        private void ImportGongFa(ISheet sheet, CgroupImportResult result)
        {
            int orderNo = 0;
            GongFa current = null;
            string currentGroup = null; // 当前分组：人工/辅材/机械

            for (int r = 0; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;

                string a = GetString(row, 0);
                string b = GetString(row, 1);

                // 三标签表头：A="项目编号" → 新工法
                if (a.Trim() == "项目编号")
                {
                    var gf = new GongFa
                    {
                        Code = GetString(row, 1),
                        Name = GetString(row, 5),
                        Unit = GetString(row, 8),
                        Part = "通用"
                    };
                    gf.Id = _repo.InsertGongFa(gf);
                    result.GongFaIds.Add(gf.Id);
                    result.GongFaCount++;
                    current = gf;
                    orderNo = 0;
                    currentGroup = null;
                    continue;
                }
                if (current == null) continue;

                // 小计/合计/管理费/利润行：跳过，不影响分组
                if (b.Contains("小计") || b.Contains("合计") || b.Contains("管理费") || b.Contains("利润") || b.Contains("综合单价"))
                    continue;

                // B列出现分组名 → 更新当前分组
                if (b.Contains("人工")) currentGroup = "人工";
                else if (b.Contains("机械")) currentGroup = "机械";
                else if (b.Contains("设备")) currentGroup = "设备";
                else if (b.Contains("材料")) currentGroup = "辅材";

                string resName = GetString(row, 2);
                if (string.IsNullOrWhiteSpace(resName)) continue;

                orderNo++;
                var price = GetDecimal(row, 5);
                var res = new GongFaResource
                {
                    GongFaId = current.Id,
                    OrderNo = orderNo,
                    ResourceType = currentGroup ?? "辅材",
                    Name = resName,
                    Unit = GetString(row, 3),
                    Dosage = GetDecimal(row, 4),
                    Price = price,
                    IsPlaceholder = price == null ? 1 : 0
                };
                _repo.InsertResource(res);
                result.ResourceCount++;
            }
        }

        // 主材导入模板 sheet（带价格）：类目/名称/规格/单位/收入价/消耗量
        // 既写主材主档，也作为"主材"资源挂到本工作簿对应的工法上（可追溯关联）
        private int ImportMainMaterialWithPrice(ISheet sheet, List<long> gongFaIds)
        {
            int n = 0, orderNo = 1000; // 主材资源排在工/料/机之后
            // 该输入模板一个工作簿对应一个工法；多工法且主材sheet无编号列时无法区分，取该工作簿的工法
            long? targetGongFaId = gongFaIds.Count == 1 ? gongFaIds[0] : (gongFaIds.Count > 0 ? (long?)gongFaIds[0] : null);

            for (int r = 1; r <= sheet.LastRowNum; r++) // 跳过表头
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;
                string name = GetString(row, 1);
                if (string.IsNullOrWhiteSpace(name)) continue;

                string spec = GetString(row, 2);
                string unit = GetString(row, 3);
                decimal? incomePrice = GetDecimal(row, 4);
                decimal? dosage = GetDecimal(row, 5);

                _repo.UpsertMainMaterial(new MainMaterial
                {
                    Name = name,
                    Spec = spec,
                    Unit = unit,
                    IncomePrice = incomePrice,
                    Dosage = dosage
                });

                // 挂为工法主材资源（Price=收入价，Dosage=消耗量）
                if (targetGongFaId.HasValue)
                {
                    _repo.InsertResource(new GongFaResource
                    {
                        GongFaId = targetGongFaId.Value,
                        OrderNo = orderNo++,
                        ResourceType = "主材",
                        Name = name,
                        Spec = spec,
                        Unit = unit,
                        Dosage = dosage,
                        Price = incomePrice,
                        IsPlaceholder = incomePrice == null ? 1 : 0
                    });
                }
                n++;
            }
            return n;
        }

        // 模板1主材目录（12列，无价格）：名称在第7列，规格第8列，单位第10列；仅补缺
        private int ImportMainMaterialCatalog(ISheet sheet)
        {
            int n = 0;
            for (int r = 2; r <= sheet.LastRowNum; r++) // 跳过R1标题、R2表头
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;
                string name = GetString(row, 6);  // G列=名称
                if (string.IsNullOrWhiteSpace(name)) continue;

                var mat = new MainMaterial
                {
                    Name = name,
                    Spec = GetString(row, 7),
                    Unit = GetString(row, 9)
                };
                if (_repo.InsertIfNotExistsMainMaterial(mat)) n++;
            }
            return n;
        }

        // 管理费利润计算公式 sheet
        private int ImportFeeRules(ISheet sheet)
        {
            int n = 0;
            for (int r = 1; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;
                string name = GetString(row, 0);
                if (string.IsNullOrWhiteSpace(name) || name.Contains("注")) continue;
                _repo.InsertFeeRule(new FeeRule
                {
                    Name = name,
                    Base = GetString(row, 1),
                    Rate = GetDecimal(row, 2)
                });
                n++;
            }
            return n;
        }

        // ---- NPOI 取值帮助 ----
        private string GetString(IRow row, int idx)
        {
            if (row == null) return "";
            var cell = row.GetCell(idx);
            if (cell == null) return "";
            switch (cell.CellType)
            {
                case CellType.String: return cell.StringCellValue.Trim();
                case CellType.Numeric: return cell.NumericCellValue.ToString();
                case CellType.Boolean: return cell.BooleanCellValue.ToString();
                case CellType.Formula:
                    try { return cell.StringCellValue?.Trim() ?? ""; }
                    catch { return cell.NumericCellValue.ToString(); }
                default: return "";
            }
        }

        private decimal? GetDecimal(IRow row, int idx)
        {
            if (row == null) return null;
            var cell = row.GetCell(idx);
            if (cell == null) return null;
            try
            {
                switch (cell.CellType)
                {
                    case CellType.Numeric: return (decimal)cell.NumericCellValue;
                    case CellType.String:
                        decimal d;
                        if (decimal.TryParse(cell.StringCellValue.Trim(), out d)) return d;
                        return null;
                    case CellType.Formula: return (decimal)cell.NumericCellValue;
                    default: return null;
                }
            }
            catch { return null; }
        }
    }
}
