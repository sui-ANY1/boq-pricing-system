using System.Collections.Generic;
using Microsoft.Data.Sqlite;
using WebApplication1.Models;
using WebApplication1.Repositories;

namespace WebApplication1.Services
{
    /// <summary>
    /// 其他费用 应用层：编排业务规则（校验、查重），调用仓储。
    /// 返回 (结果, 错误信息)。
    /// </summary>
    public class OtherFeeService
    {
        private readonly OtherFeeRepository _repo = new OtherFeeRepository();

        /// <summary>新增：返回新Id与错误信息</summary>
        public (long id, List<string> errors) Add(OtherFee fee)
        {
            var errors = new List<string>();
            if (fee == null) { errors.Add("请求体不能为空"); return (0, errors); }

            errors = fee.Validate();
            if (_repo.ExistsByName(fee.ProjectId, fee.Name))
                errors.Add("同一项目下工项名称已存在");

            if (errors.Count > 0) return (0, errors);

            // 先查重再插入之间存在并发窗口(TOCTOU)，直接兜底捕获唯一约束冲突
            try
            {
                long id = _repo.Insert(fee);
                return (id, errors);
            }
            catch (SqliteException ex) when (IsUniqueConstraint(ex))
            {
                errors.Add("同一项目下工项名称已存在（并发冲突）");
                return (0, errors);
            }
        }

        public List<OtherFee> List(string projectId)
        {
            return _repo.ListByProject(projectId);
        }

        /// <summary>修改：返回错误信息列表，空表示成功</summary>
        public List<string> Update(long id, OtherFee fee)
        {
            if (fee == null) return new List<string> { "请求体不能为空" };
            fee.Id = id;

            var errors = fee.Validate();
            var old = _repo.GetById(id);
            if (old == null) { errors.Add("记录不存在"); return errors; }

            fee.ProjectId = old.ProjectId; // 项目不允许改，先对齐再查重，保证查重与落库一致
            if (_repo.ExistsByName(fee.ProjectId, fee.Name, id))
                errors.Add("同一项目下工项名称已存在");

            if (errors.Count > 0) return errors;

            try
            {
                _repo.Update(fee);
            }
            catch (SqliteException ex) when (IsUniqueConstraint(ex))
            {
                errors.Add("同一项目下工项名称已存在（并发冲突）");
            }
            return errors;
        }

        public bool Delete(long id)
        {
            if (_repo.GetById(id) == null) return false;
            _repo.Delete(id);
            return true;
        }

        /// <summary>判断是否 SQLite UNIQUE 约束冲突（错误码19，扩展码2067）</summary>
        private static bool IsUniqueConstraint(SqliteException ex)
        {
            return ex.SqliteExtendedErrorCode == 2067 || ex.SqliteErrorCode == 19;
        }
    }
}
