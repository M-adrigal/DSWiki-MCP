# DSWiki-MCP

从 Word / PPT 提取排版规格、章节结构和文风硬指标，供 Agent 建立文档画像、验收成稿、判断一批文档是否属于同一套模板。

服务只输出可复现的数字（字号、缩进、编号、句长、技术词密度等）。规则命名、内容套路、风格描述由调用方完成。评分权重是经验值，用于同一份文档的纵向比较；交付检查（页码、页眉、表题、修订记录）反映完成度，不计分。

## 安装

Python 3.10+，stdio 运行，由 MCP 客户端拉起子进程，不需要起 HTTP 服务。支持 `.docx` 和 `.pptx`，按扩展名分流。

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

在客户端的 MCP 配置里加一项，路径换成实际位置：

```json
{
  "mcpServers": {
    "DSWiki-MCP": {
      "command": "/绝对路径/到/.venv/bin/python3",
      "args": ["/绝对路径/到/DSWiki-MCP/mcp_server/server.py"],
      "timeout": 120000
    }
  }
}
```

配置文件位置因客户端而异，WorkBuddy 是 `~/.workbuddy/mcp.json`。改完需重启客户端。也可以前台运行 `python3 mcp_server/server.py` 看输出。

数据目录默认 `~/.workbuddy/doc-style/`，用 `DOC_STYLE_HOME` 覆盖，首次启动自动创建。画像不随代码走，换工作区或重装仓库不会丢。

依赖只有标准库和 `requirements.txt` 里的四个包，无数据库、无网络调用、无密钥。

## 目录

```text
mcp_server/server.py     MCP 入口，9 个工具 + 运维 + 画像生命周期
docx_probe.py            Word 排版与结构
pptx_probe.py            PPT 视觉规格
style_probe.py           文风硬指标
quality_gate.py          Word 质量门与交付检查
pptx_quality_gate.py     PPT 质量门（首次调用时加载）
cluster_gate.py          十维特征聚类
batch_compare.py         按聚类组比对规格
profile_doctor.py        画像库体检（可命令行运行）
service_config.py        判定阈值
version.py               口径版本、CHANGELOG、指纹、重算判据
```

数据目录不在仓库内，结构如下：

```text
$DOC_STYLE_HOME/
├── config.json             判定阈值，缺省时用 service_config.DEFAULTS
├── profiles/<profile_id>/  DNA.md / rules.json / spec.json / meta.json
├── audit.jsonl             变更审计：谁改了什么
└── calls.jsonl             运行指标：工具、耗时、结果状态
```

## 工具

以 `@mcp.tool()` 实际注册的 9 个为准。`check_profiles`、`compare_documents`、`list_profiles` 是内部函数，不对外暴露。

| 工具 | 用途 | 关键参数 |
| --- | --- | --- |
| `analyze_document` | 单份文档的质量门与硬指标 | `path`、`preset`、`doc_type`、`want`、`detail` |
| `analyze_batch` | 一批文档筛选、聚类、组内比对 | `paths`、`doc_type`、`with_compare`、`exclude_rejected` |
| `get_profile` | 读一份画像，不传 id 则列清单 | `profile_id`、`track` |
| `save_profile` | 画像落库，逐条校验规则证据 | `profile_id`、`dna_markdown`、`rules`、`meta`、`overwrite` |
| `edit_rules` | 增改删单条规则 | `action`、`rule_id`、`patch`、`reason` |
| `manage_profile` | 状态与元信息 | `action`、`confirmed_by`、`patch` |
| `merge_profiles` | 合并拆错的画像 | `target_id`、`source_ids`、`reason` |
| `backup_profiles` | 画像库打包留底 | `out_path` |
| `service_control` | 状态、指标、配置、重载、备份 | `action`、`patch`、`confirm` |

常见症状：技术文档分数偏低，多半没传 `doc_type`；批量报「可用文档少于 2 份」，先在「已跳过」里找原因；读一次画像使用次数就增加，用 `track=False`。

## 使用约定

`doc_type` 决定按哪套标准判，默认 `auto`（通用 / 公文口径，要求首行缩进 2 字符、必须有目录）。技术文档不缩进是常态、段落少于 150 时缺目录不扣分，批量分析必须显式传 `doc_type=技术文档`，否则会被判成拒收、不参与建画像，且与单文档入口的结论不一致。批量返回里的「质量门口径」写明本次用的标准。可选值：`auto` / `公文` / `技术文档` / `网页` / `报告`。

`preset` 决定一次拿多少内容：

| preset | 用途 | 体积 |
| --- | --- | --- |
| `audit` | 验收成稿：质量门 + 交付检查 + 排版提示 | 约 1–2 KB |
| `build_profile` | 建画像：质量门 + 文风 + 章节结构 | 约 10 KB |
| `full` | 全量，含表格与样式明细，排查用 | 约 14 KB |

`want` 可进一步裁剪（Word：quality / format / voice；PPT：quality / style）。返回里出现「先读这个」表示文风指标有自检告警，先读告警说明，不要拿那个数字写规则。

写操作有确认门槛：`service_control(action=set)` 不带 `confirm=True` 只预览；`manage_profile(action=promote)` 必须给 `confirmed_by`；覆盖「生效」画像会被拒。`save_profile` 的 `meta` 必须带「文档类型」，它是按类型检索和重复检测的依据。

## 口径与版本

口径分四层，改东西前先对号入座：

| 层 | 由谁决定 | 变了会怎样 |
| --- | --- | --- |
| 统计口径 | 6 个统计引擎模块 | 已有画像数字作废，需全量重算 |
| 判定口径 | `doc_type` 入参 | 版本不变，同一份文档分数可能不同 |
| 准入阈值 | `config.json` | 版本不变，样本构成会变 |
| 呈现口径 | `preset` / `detail` | 只影响返回体积 |

每次分析返回都带 `_审计`：口径版本、规则指纹（6 个统计模块）、工具层指纹（`server.py` + `service_config.py`）、输入指纹。规则指纹对引擎模块的任何改动都会变，包括只改展示措辞。指纹变了按三步确认：比输入指纹，再比同一份文档的 `score`，分数变了才需要重算。

要不要重算看 `version.py` 里 CHANGELOG 的「需重算」字段，不看版本号，也不单看指纹。当前版本 `2026.11.10.1`，统计口径未改，`需重算=False`。改统计口径时升版本、记变更与影响指标、全量重算；只改接口或返回结构时升版本并标 `需重算=False`。

## 配置与运维

判定阈值用 `service_control(action=config)` 读、`action=set` 改（`confirm=True` 才落盘），默认值在 `service_config.DEFAULTS`。这里只放阈值，统计口径改动必须走版本升级和全量重算。

`service_control(action=status)` 给出各引擎模块是否需要重载；改了 `style_probe`、`quality_gate`、`cluster_gate` 等模块后用 `action=reload` 热加载，`server.py` 自身改动和新增工具都要重启客户端。运行指标看 `action=stats`：成功率按 `ok / (ok + bug)` 算，`bad_input` 是调用方传错、`rejected` 是前置条件不满足，两者另列比率。`audit.jsonl` 记变更，`calls.jsonl` 记调用，别混用。

画像库体检可运行 `python3 profile_doctor.py`。

## 已知边界

- 手写的「目 录」加目录项不算目录，需目录域或标题样式，规范样本可能因此扣 5 分。
- 批量入口对非法 `doc_type` 没有统一校验，逐份报「无法解析」后顶层错误变成「可用文档少于 2 份」，单文档入口则会直接拒绝并列出合法取值。
- `exclude_rejected=false` 时「本次用哪套标准」写「未启用」，下方说明文字仍是默认口径模板。
- 文件头注释、`INSTRUCTIONS` 和 `_TOOL_NAMES` 列了未注册的 `check_profiles`、漏了 `backup_profiles`，属说明文字滞后。
- PPT 链路不做文风量化。

## 仓库维护

`.gitignore` 已覆盖缓存、系统文件、虚拟环境、样本文档和数据文件；测试脚本与自检画像已移除。数据目录不要纳入版本库，画像的 `spec.json` 含本地文件路径。

许可 MIT，见 `LICENSE`。本仓库为私有仓库。
