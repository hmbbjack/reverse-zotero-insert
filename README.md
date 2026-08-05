# 逆向 Zotero 插入（Reverse Zotero Field-Code Insertion）

在**无法操作 Word 的 Zotero 插件 GUI** 的场景下（命令行 / 自动化 / Claude Code skill），
把一份 `.docx` 里的纯文字参考文献（正文引文 + 文末参考文献表）转换成 **Zotero 能直接
识别、管理、刷新的域代码（field code）**。效果与用 Zotero 插件插入一致——在 Word 里用
Zotero 打开时，引文和参考文献表都是"活的"。

本仓库同时提供**前置引用摄取**：从纯文字编号参考文献出发，联网解析元数据 → 导入 Zotero
→ 再插入域代码，形成端到端工作流。

## 特性

- **端到端**：纯文字编号参考文献 → 元数据解析（CrossRef/PubMed 搜索 API）→ 导入 Zotero
  → 域代码插入 → 校验。
- **前置 MCP 配置检查**：工作流第一步自动探测 Zotero MCP 是否就绪，未就绪时提醒用户配置。
- **强制用户检查点**：导入前向用户展示元数据汇总，并询问目标文件夹（列出现有 collection
  或新建），未获用户确认**绝不写入** Zotero。
- **导入前去重**：按标题/DOI 检查库中是否已有相同条目，供用户选择跳过/新建/复用。
- **逆向域代码格式**：直接改写 `word/document.xml`，其余 zip 部分原样保留。

## 前置依赖

本工具依赖 [Zotero](https://www.zotero.org/) 及其 MCP 服务，请先完成以下安装：

1. **Zotero**（7.0+）：从 [zotero.org](https://www.zotero.org/) 下载安装。
2. **Zotero MCP 插件**：基于 [cookjohn/zotero-mcp](https://github.com/cookjohn/zotero-mcp)（MIT，向作者 @cookjohn 致敬）。从其 [Releases](https://github.com/cookjohn/zotero-mcp/releases) 下载 `zotero-mcp-plugin-x.x.x.xpi`，在 Zotero 中 `工具 → 附加组件` 安装并重启，然后在 `首选项 → Zotero MCP Plugin` 中启用服务（默认端口 `23120`）。

> 本仓库与 cookjohn/zotero-mcp 无隶属关系，仅作为下游使用者致谢。

## 安装

```bash
pip install -r requirements.txt
# 需要能访问本机 Zotero MCP（见下方"配置"）
```

## 配置

代码中**不含任何硬编码的个人标识**。所有个人值通过环境变量或派生获得：

| 环境变量 | 说明 |
|----------|------|
| `ZOTERO_USER_ID` | Zotero 用户 ID（未设时自动从 `~/Zotero/zotero.sqlite` 的 `users` 表推导） |
| `ZOTERO_URI_PREFIX` | 条目 URI 前缀，默认 `http://zotero.org/users/<userID>/items/` |
| `ZOTERO_MCP_URL` | Zotero MCP 服务地址，默认 `http://127.0.0.1:23120/mcp` |
| `CROSSREF_MAILTO` | CrossRef 礼貌池标识（可选，推荐填写以降低限流） |
| `NCBI_API_KEY` | NCBI eutils API key（可选） |

示例（`config.example.env`，复制为 `config.local.env` 即可被自动加载）：

```
ZOTERO_USER_ID=CHANGE_ME
CROSSREF_MAILTO=CHANGE_ME
```

或直接导出环境变量。

### Zotero MCP

本工具通过本机 Zotero MCP 服务（Streamable HTTP，默认 `127.0.0.1:23120/mcp`）写入条目。
安装方式见上方[前置依赖](#前置依赖)。请确保 Zotero 正在运行且已启用 MCP 连接器，工作流开始时会自动探测，未就绪会提示。

## 工作流

1. **前置 MCP 检查**：`check_zotero_mcp()` 探测 MCP 是否就绪；未就绪则提示用户配置。
2. **解析参考文献**：`extract_reference_paragraphs()` 从 docx 读出编号参考文献。
3. **批量解析元数据**：`resolve_metadata()` 用 CrossRef/PubMed 搜索 API 核实，产出
   `resolution_summary.json`（含置信度）。
4. **去重检查**：`dedup_check()` 用 `search_library` 查重，产出 `dedup_report.json`。
5. **强制检查点**：展示汇总表 + 询问目标文件夹（现有 collection 或新建），以及每篇重复的处置。
6. **导入**：`import_to_zotero()` 写入 Zotero 到指定文件夹。
7. **构建 mapping**：`build_item_mapping()` 产出 `item_mapping.json`。
8. **插入域代码**：`zotero_field_insert.insert_zotero_fields()` + `verify()`。

```python
from reference_ingest import extract_reference_paragraphs, resolve_metadata, dedup_check, import_to_zotero, build_item_mapping
from zotero_field_insert import insert_zotero_fields, verify

refs = extract_reference_paragraphs("论文.docx")
resolved = resolve_metadata(refs)          # resolution_summary.json
dup = dedup_check(resolved)                # dedup_report.json
# —— 此处为强制检查点：向用户展示 + 确认目标文件夹/去重处置 ——
plan = {k: {"action": "import", "csl": v["csl"]} for k, v in resolved.items() if v["csl"]}
imported = import_to_zotero(plan, "COLLECTION_KEY")
mapping = build_item_mapping(imported)     # item_mapping.json
insert_zotero_fields("论文.docx", "论文_zotero.docx", (1, 118), list(range(119, 180)), mapping)
print(verify("论文_zotero.docx", [m["itemKey"] for m in mapping.values()]))
```

## 校验

`verify(out, valid_keys)` 检查：fldChar begin/separate/end 配对平衡、所有 JSON 合法、所有
URI 指向真实库条目、itemData 含 type/title/author/issued。若传入 `src_docx` 与
`item_mapping`,还会比对原文做**引文零遗漏 + 显示文本逐字一致**校验（在原文里用引文匹配器
找出所有"作者+年份",逐一核对是否都出现在输出域的显示文本中）。

## 目录

```
├── SKILL.md                 # skill 触发 + 端到端工作流 + 检查点规则
├── reference_ingest.py      # 前置引用摄取（解析/解析元数据/去重/导入/mapping）
├── zotero_field_insert.py   # 域代码插入器（引文匹配 + 域构造 + 校验）
├── zotero_mcp.py            # Zotero MCP 客户端（Streamable HTTP）
├── config.py                # 配置（env + 自动推导，无硬编码 ID）
└── tests/                   # 单元测试
```

## 许可

MIT。详见 [LICENSE](LICENSE)。