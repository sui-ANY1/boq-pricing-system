using System;

namespace WebApplication1.Models
{
    /// <summary>
    /// 其他费用 领域实体
    /// 业务规则内聚在此层：利润、利润率由实体自己算，校验也在实体层。
    /// </summary>
    public class OtherFee
    {
        /// <summary>记录标识（自增主键）</summary>
        public long Id { get; set; }

        /// <summary>所属项目标识</summary>
        public string ProjectId { get; set; }

        /// <summary>工项名称（必填，同一项目内唯一）</summary>
        public string Name { get; set; }

        /// <summary>收入价（元，可为空，空按0处理）</summary>
        public decimal? IncomePrice { get; set; }

        /// <summary>成本价（元，必填，必须 >= 0）</summary>
        public decimal CostPrice { get; set; }

        /// <summary>备注（可为空）</summary>
        public string Remark { get; set; }

        /// <summary>收入实际值：空按0，保留两位小数</summary>
        public decimal Income
        {
            get { return Math.Round(IncomePrice ?? 0m, 2); }
        }

        /// <summary>利润 = 收入 - 成本（服务端计算，不允许前端传）</summary>
        public decimal Profit
        {
            get { return Math.Round(Income - CostPrice, 2); }
        }

        /// <summary>利润率 = 利润 / 收入（收入为0时按0），保留四位小数（展示为百分比时对应两位小数）</summary>
        public decimal ProfitRate
        {
            get
            {
                if (Income == 0) return 0m;
                return Math.Round(Profit / Income, 4); // 存小数，展示时按百分比
            }
        }

        /// <summary>
        /// 领域校验：返回错误信息列表，空表示通过。
        /// </summary>
        public System.Collections.Generic.List<string> Validate()
        {
            var errors = new System.Collections.Generic.List<string>();

            if (string.IsNullOrWhiteSpace(ProjectId))
                errors.Add("项目Id不能为空");

            if (string.IsNullOrWhiteSpace(Name))
                errors.Add("工项名称不能为空");

            if (CostPrice < 0)
                errors.Add("成本价不能为负数");

            if (IncomePrice.HasValue && IncomePrice.Value < 0)
                errors.Add("收入价不能为负数");

            return errors;
        }
    }
}
