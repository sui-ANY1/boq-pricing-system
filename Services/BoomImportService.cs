using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;
using WebApplication1.Models;
using WebApplication1.Repositories;

namespace WebApplication1.Services
{
    /// <summary>
    /// 分部分项导入：读 分部分项清单.json，写入 BoomItem 表。
    /// 工程量=parts[].quantity 求和；收入主材=income_price、收入综合单价=income.price、施工=综合单价-主材；
    /// 成本按需求2.4：成本主材=收入主材，成本施工=收入施工×80%。
    /// </summary>
    public class BoomImportService
    {
        private readonly CgroupRepository _repo = new CgroupRepository();

        public int Import(string jsonPath)
        {
            _repo.ClearBoom();

            string text = File.ReadAllText(jsonPath, System.Text.Encoding.UTF8);
            JObject root = JObject.Parse(text);
            var schemes = root["schemes"] as JArray;

            int count = 0;
            foreach (var s in schemes)
            {
                string name = (string)s["name"] ?? "";
                string unit = (string)s["unit"] ?? "";

                var materials = s["materials"] as JArray;
                if (materials == null || materials.Count == 0) continue;

                var m0 = materials[0];
                var mat = m0["material"];

                // 工程量 = parts[].quantity 求和
                decimal qty = 0;
                var parts = m0["parts"] as JArray;
                if (parts != null)
                    foreach (var pt in parts)
                        qty += (decimal?)pt["quantity"] ?? 0;

                // 收入侧：主材费=income_price，综合单价=income.price，施工费=综合单价-主材费
                decimal incMaterial = (decimal?)mat?["income_price"] ?? 0;
                decimal incTotal = (decimal?)mat?["income"]?["price"] ?? 0;
                decimal incConstruct = incTotal - incMaterial;

                // 成本侧按需求 2.4（本文件为准，不采用模型原始成本）：
                // 成本主材=收入主材；成本施工=收入施工×80%；成本小计=两者之和
                decimal costMaterial = incMaterial;
                decimal costConstruct = Math.Round(incConstruct * 0.8m, 2);
                decimal costTotal = Math.Round(costMaterial + costConstruct, 2);

                var item = new BoomItem
                {
                    GongFaCode = (string)mat?["income"]?["code"] ?? "",
                    Name = name,
                    MaterialName = (string)mat?["name"] ?? "",
                    Spec = (string)mat?["spec"] ?? "",
                    Unit = unit,
                    Quantity = qty,
                    IncomeMaterial = incMaterial,
                    IncomeConstruct = incConstruct,
                    IncomeTotal = incTotal,
                    CostMaterial = costMaterial,
                    CostConstruct = costConstruct,
                    CostTotal = costTotal
                };
                _repo.InsertBoomItem(item);
                count++;
            }
            return count;
        }
    }
}
