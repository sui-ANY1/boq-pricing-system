using System;

namespace WebApplication1.Models
{
    /// <summary>工法主表（对应 A1 第7题主表）</summary>
    public class GongFa
    {
        public long Id { get; set; }
        public string Code { get; set; }      // 项目编码
        public string Name { get; set; }       // 工法名称
        public string Unit { get; set; }      // 计量单位
        public string Part { get; set; }       // 部位
    }

    /// <summary>工法资源明细（主表的一对多明细：人工/辅材/机械）</summary>
    public class GongFaResource
    {
        public long Id { get; set; }
        public long GongFaId { get; set; }
        public int OrderNo { get; set; }
        public string ResourceType { get; set; } // 人工/辅材/机械/主材
        public string Name { get; set; }
        public string Spec { get; set; }
        public string Unit { get; set; }
        public decimal? Dosage { get; set; }    // 含量/消耗量
        public decimal? Price { get; set; }     // 单价，空=占位行
        public int IsPlaceholder { get; set; }
    }

    /// <summary>主材主档（模板1：名称+收入价+消耗量）</summary>
    public class MainMaterial
    {
        public long Id { get; set; }
        public string Name { get; set; }
        public string Spec { get; set; }
        public string Unit { get; set; }
        public decimal? IncomePrice { get; set; } // 收入价
        public decimal? Dosage { get; set; }      // 消耗量
    }

    /// <summary>取费规则（管理费/利润）</summary>
    public class FeeRule
    {
        public long Id { get; set; }
        public string Name { get; set; }
        public string Base { get; set; }   // 基数表达式
        public decimal? Rate { get; set; } // 费率
    }

    /// <summary>分部分项清单行（扁平行+路径串）</summary>
    public class BoomItem
    {
        public long Id { get; set; }
        public string Path { get; set; }
        public string GongFaCode { get; set; }
        public string Name { get; set; }
        public string MaterialName { get; set; }
        public string Spec { get; set; }
        public string Unit { get; set; }
        public decimal? Quantity { get; set; }
        public decimal? IncomeMaterial { get; set; }
        public decimal? IncomeConstruct { get; set; }
        public decimal? IncomeTotal { get; set; }
        public decimal? CostMaterial { get; set; }
        public decimal? CostConstruct { get; set; }
        public decimal? CostTotal { get; set; }
    }
}
