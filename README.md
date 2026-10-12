# 工程造价清单管理系统（WebApplication1）

基于 **ASP.NET MVC 5 + Web API** 的工程造价清单后端服务，围绕「其他费用」「工法组价」「分部分项清单」三大业务，提供数据录入、Excel / JSON 导入、自动算量组价与 Excel 导出，数据统一落地到 **SQLite 单文件数据库**，零外部数据库依赖。

---

## 功能特性

- **其他费用（B 组）**：按项目对工项进行增删改查；服务端自动计算利润、利润率；支持 Excel 批量导入与清单导出；首次启动自动初始化 6 个固定工项。
- **工法组价（C 组）**：读取「定额计价工法」「主材」两个 Excel 模板，解析工法主表及其人工 / 辅材 / 机械 / 主材 / 设备资源明细；依据取费规则（管理费、利润，按「基数表达式 × 费率」）自动计算综合单价。
- **分部分项清单**：从 JSON 导入清单，工程量按 `parts.quantity` 汇总；测算收入与成本（成本施工费 = 收入施工费 × 80%）；导出分部分项清单 Excel。
- **工法主材汇总表**：按组价结果导出汇总 Excel。
- **零配置存储**：SQLite 在应用启动时幂等建表；提供 Reset 接口一键清库、重读模板、重新生成导出表，运行结果不依赖历史残留数据。
- 内置 **Web API Help Page** 在线接口文档。

---

## 技术栈

| 类别 | 选型 |
| --- | --- |
| 运行框架 | .NET Framework 4.8 |
| Web 框架 | ASP.NET MVC 5.2.9、ASP.NET Web API 5.2.9 |
| 数据存储 | SQLite（Microsoft.Data.Sqlite 10 + SQLitePCLRaw / e_sqlite3） |
| Excel 处理 | NPOI 2.8 |
| JSON | Newtonsoft.Json 13.0.3 |
| 数值计算 | MathNet.Numerics、ExtendedNumerics.BigDecimal |
| 前端资源 | Bootstrap 5.2.3、jQuery 3.7.0、Modernizr 2.8.3 |
| 开发 / 托管 | Visual Studio 2019/2022、IIS Express |
| 包管理 | NuGet（packages.config，旧式引用） |

---

## 目录结构

```
WebApplication1/
├── App_Start/                 # 路由 / 过滤器 / 打包 配置
│   ├── BundleConfig.cs        # 静态资源打包
│   ├── FilterConfig.cs        # 全局过滤器
│   ├── RouteConfig.cs         # MVC 路由
│   └── WebApiConfig.cs        # Web API 路由（特性路由 + 默认路由）
├── Areas/HelpPage/            # Web API 帮助页（访问 /Help）
├── Content/                   # Bootstrap 5、Site.css
├── Controllers/
│   ├── CgroupController.cs    # C 组：工法组价 + 分部分项 导入/导出/重置
│   ├── OtherFeeController.cs  # B 组：其他费用 CRUD / 导入 / 导出 / 重置
│   ├── SqliteTestController.cs# SQLite 连通性测试
│   ├── HomeController.cs      # MVC 首页
│   └── ValuesController.cs    # Web API 模板示例
├── Models/
│   ├── CgroupModels.cs        # GongFa / GongFaResource / MainMaterial / FeeRule / BoomItem
│   └── OtherFee.cs            # 其他费用实体（内聚校验与利润计算）
├── Repositories/
│   ├── Db.cs                  # 建表 / 固定工项初始化 / 重置
│   ├── CgroupRepository.cs    # C 组数据访问
│   └── OtherFeeRepository.cs  # 其他费用数据访问
├── Services/
│   ├── CgroupImportService.cs # 工法 / 主材模板导入
│   ├── CgroupCalcService.cs   # 工法组价计算
│   ├── CgroupExportService.cs # 工法主材汇总表导出
│   ├── BoomImportService.cs   # 分部分项 JSON 导入
│   ├── BoomExportService.cs   # 分部分项清单导出
│   ├── OtherFeeService.cs     # 其他费用业务编排（校验 / 查重）
│   ├── OtherFeeImportService.cs
│   └── OtherFeeExportService.cs
├── Scripts/                   # jQuery、Bootstrap JS
├── App_Data/                  # 样例数据（test.db 运行时生成，已忽略）
├── Global.asax / Global.asax.cs  # 应用启动入口（启动时 Db.Init()）
├── Web.config                 # 应用配置（含 ResetKey、程序集绑定重定向）
├── packages.config            # NuGet 依赖清单
├── e_sqlite3.dll              # SQLite 原生库（生成时复制到输出目录）
├── 分部分项清单.json           # 分部分项导入数据源
└── *.xlsx                     # 各类输入 / 参考模板（见下方“数据与模板说明”）
```

> 分层约定：`Controller（接口） → Service（业务编排 / 计算） → Repository（数据访问） → Db（SQLite 基础设施）`；领域校验与派生指标（如利润）内聚在 `Model`。

---

## 数据表结构

应用启动时由 `Db.Init()` 幂等创建以下 6 张表：

| 表 | 说明 |
| --- | --- |
| `OtherFee` | 其他费用（Id、ProjectId、Name、IncomePrice、CostPrice、Remark），`(ProjectId,Name)` 唯一 |
| `GongFa` | 工法主表（编码、名称、计量单位、部位） |
| `GongFaResource` | 工法资源明细（人工 / 辅材 / 机械 / 主材 / 设备、含量、单价、是否占位行） |
| `MainMaterial` | 主材主档（名称、规格、单位、收入价、消耗量） |
| `FeeRule` | 取费规则（名称、基数表达式、费率） |
| `BoomItem` | 分部分项清单行（路径、工法编码、名称、工程量、收入 / 成本的主材 / 施工 / 合计） |

首次启动会向默认项目（`ProjectId = DEFAULT`）插入 6 个固定工项：岩土工程、建筑工程、技术措施费、安全文明施工及其他措施项目、规费、其它工程。

---

## HTTP API

### 其他费用 OtherFee

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| POST | `/api/otherfee` | 新增工项（请求体为 OtherFee JSON） |
| GET | `/api/otherfee?projectId=xxx` | 查询某项目的全部工项 |
| PUT | `/api/otherfee/{id}` | 修改工项 |
| DELETE | `/api/otherfee/{id}` | 删除工项 |
| POST | `/api/otherfee/import?projectId=xxx` | 上传 Excel（multipart/form-data）批量导入 |
| GET | `/api/otherfee/export?projectId=xxx&projectName=yyy` | 导出「其他费用清单.xlsx」 |
| POST | `/api/otherfee/reset?key=xxx` | 清库重建 |

### 工法组价 / 分部分项 Cgroup

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| POST | `/api/cgroup/import` | 读取工法 / 主材模板与分部分项 JSON，写入数据库 |
| GET | `/api/cgroup/price` | 查看全部工法组价结果 |
| GET | `/api/cgroup/export` | 导出「工法主材汇总表.xlsx」 |
| POST | `/api/cgroup/boomimport` | 单独导入分部分项 JSON |
| GET | `/api/cgroup/boom` | 查看分部分项数据 |
| GET | `/api/cgroup/boomexport` | 导出「分部分项清单.xlsx」 |
| POST | `/api/cgroup/reset?key=xxx` | 清 C 组表 → 重读模板 → 重新生成两张导出表 |

### 其他

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| GET | `/api/sqlitetest/testconnect` | 测试 SQLite 连接与驱动 |
| GET | `/help` | Web API 在线帮助文档 |
| GET | `/` | MVC 首页 |
| GET/POST | `/api/values` | Web API 模板自带示例 |

---

## 快速开始

1. **环境准备**：安装 Visual Studio 2022（勾选「ASP.NET 和 Web 开发」工作负载，含 .NET Framework 4.8 开发包与 IIS Express）。
2. **打开项目**：直接打开 `WebApplication1.csproj`（仓库未附带 `.sln`，VS 打开后可自行另存解决方案）。
3. **还原 NuGet 包**：生成时会自动还原；也可右键项目 / 解决方案 →「还原 NuGet 包」。
4. **运行**：按 `F5` 以 IIS Express 启动。
5. **访问**：首页 `/`，接口文档 `/Help`。
6. **典型调用流程（C 组）**：
   - `POST /api/cgroup/import`：导入工法、主材及分部分项；
   - `GET /api/cgroup/price`：查看组价结果；
   - `GET /api/cgroup/export`：下载工法主材汇总表；
   - `GET /api/cgroup/boomexport`：下载分部分项清单。

> 根目录的输入模板（xlsx / json）在 csproj 中配置为 `CopyToOutputDirectory = PreserveNewest`，生成时会自动复制到运行目录（bin），接口从运行目录读取，无需手工拷贝。

---

## 计算口径

**其他费用**

- 收入 = 收入价（为空按 0，保留两位小数）
- 利润 = 收入 − 成本（服务端计算，不接受前端传入）
- 利润率 = 利润 / 收入（收入为 0 时记 0）

**工法组价**

- 资源金额 = 含量 × 有效单价（单价为空的占位 / 缺价资源只保留业务信息、不计金额）
- 主材费 = 本工法挂接主材资源之和
- 材料费 = 主材费 + 辅材费 + 设备费
- 直接费 = 人工费 + 材料费 + 机械费
- 管理费 / 利润 = 取费规则基数表达式（将「分部分项人工费 / 材料费 / 机械费 / 主材费 / 设备费 / 直接费 / 管理费」替换为数值后做四则运算）× 费率
- 综合单价 = 直接费 + 管理费 + 利润

**分部分项成本**

- 工程量 = 各 `parts.quantity` 之和
- 收入主材费 = `income_price`；收入综合单价 = `income.price`；收入施工费 = 综合单价 − 主材费
- 成本主材费 = 收入主材费；成本施工费 = 收入施工费 × 80%；成本合计 = 两者之和

---

## 数据与模板说明

| 文件 | 用途 | 是否入库 |
| --- | --- | --- |
| `工法清单模板--1.主材导入.xlsx` | 主材主档输入（名称 / 收入价 / 消耗量） | 是（输入模板） |
| `工法清单模板--2.定额计价工法导入.xlsx` | 工法及资源明细输入 | 是（输入模板） |
| `工法清单模板--3.定额计价工法导出.xlsx` | 工法导出参考模板 | 是 |
| `建筑单体1#分部分项清单-模板.xlsx` | 分部分项导出参考模板 | 是 |
| `导出其他费用清单.xlsx` | 其他费用导入 / 参考 | 是 |
| `分部分项清单.json` | 分部分项导入数据源 | 是 |
| `分部分项清单.xlsx` | 接口导出成品 | 否（可重新生成，已忽略） |
| `工法主材汇总表.xlsx` | 接口导出成品 | 否（可重新生成，已忽略） |
| `App_Data/test.db` | 运行时数据库 | 否（自动创建，已忽略） |
| `App_Data/*.xlsx` | 其他费用导入 / 导出样例 | 是（参考样例） |

---

## 重置接口与安全说明

- `POST /api/cgroup/reset?key=reset2026`：清空 C 组全部表 → 重读模板重新导入 → 重新生成两张导出表并写回运行目录。
- `POST /api/otherfee/reset?key=reset2026`：删除并重建全部表。
- 重置密钥在 `Web.config` 的 `appSettings:ResetKey` 配置（默认 `reset2026`，`Db.cs` 内置同值兜底）。
- 该密钥仅用于本地 / 演示，**生产环境务必替换为正式身份鉴权或限制内网访问**，避免数据被任意清空。

---

## 说明

本项目为后端开发练习 / 考核用途，聚焦工程造价清单的导入、组价与导出流程，未包含完整前端页面。
