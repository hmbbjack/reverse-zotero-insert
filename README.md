# 逆向 Zotero 插入（Reverse Zotero Field-Code Insertion）· v3

在**无法操作 Word 的 Zotero 插件 GUI** 的场景下（命令行 / 自动化 / AI agent skill），
把一份 `.docx` 里的纯文字参考文献（正文引文 + 文末参考文献表）转换成 **Zotero 能直接
识别、管理、刷新的域代码（field code）**。效果与用 Zotero 插件插入一致——在 Word 里用
Zotero 打开时，引文和参考文献表都是"活的"。

本仓库同时提供**前置引用摄取**：从纯文字编号参考文献出发，联网解析元数据 → 导入 Zotero
→ 再插入域代码，形成端到端工作流。

## 维护状态

本人同时存在学业压力和科研压力，因时间原因，几乎不做小版本更新，只在想起来的时候做大版本更新。

## 特性

- **端到端**：纯文字编号参考文献 → 元数据解析（CrossRef/PubMed 搜索 API）→ 导入 Zotero
  → 域代码插入 → 逐组对账与交付前审计。
- **扫描优先（v3）**：`docx_audit.scan_document()` 一次调用给出转换前全貌——域清单、样式、
  DOI 占位组、混合段、表格引文列 vs 数据列、citationID 分布、引用条目基数。没有扫描不动手。
- **citationID 自动避让（v3）**：新域从文档既有编号 max+1 起步。v1/v2 一律从 0 编号，只要
  文档已有引文域就必然撞号（跨版本复制段落尤其明显），而 Zotero 以 citationID 索引引文，
  撞号会让刷新把两条不同引文当成同一条。
- **老式 DOI 全路径支持（v3）**：`10.1016/S0306-4530(98)00014-6` 这类内含括号的 DOI，在
  括号组、整格、库索引三处都不会被截断或漏转（`normalize_doi()` / `extract_doi_groups()`）。
- **跨段域深度感知（v3）**：Zotero 文献表域可跨上百段落。域外文本与"残留占位"扫描按文档级
  深度判定，不会把文献表里的 `https://doi.org/…` 误报成漏转。
- **库索引不静默（v3）**：`build_library_doi_index()` 返回 `(index, failures, dups)`；
  `get_item_details` 失败（多为用户已删除的条目）必须上报而不是悄悄跳过。`verify_item_keys()`
  做死链审计，`relink_item_keys()` 一键把死链重链接到库内同 DOI 的规范条目。
- **逐组对账（v3）**：`audit_output()` 从原稿重新推导占位组，与输出域逐组核对显示文本、
  DOI 集合、URI 指向——比"数域个数"强得多，能查出整组漏转与指向错条目。
- **多种正文引用风格**：除"作者+年份"外，支持 `(doi:10.x; 10.y)` / `(10.x)` 裸 DOI 组占位
  与 `citation_text_map` 精确文本括号（FDA 说明书 / NCT / EU CT 等无 DOI 灰色文献）；
  表格"整列即 DOI"用 `insert_table_citations`。
- **前置 MCP 配置检查**：工作流第一步自动探测 Zotero MCP 是否就绪，未就绪时提醒用户配置。
- **强制用户检查点**：导入前向用户展示元数据汇总，并询问目标文件夹（列出现有 collection
  或新建），未获用户确认**绝不写入** Zotero。
- **导入前去重**：按标题/DOI 检查库中是否已有相同条目，供用户选择跳过/新建/复用；DOI 匹配经
  `get_item_details` 详情确认（`search_library` 结果不含 DOI 字段）；导入幂等，导入后
  `audit_imported_by_doi()` 审计残留重复。
- **以库为准**：`apply_library_data()` 用 Zotero 库内元数据重建域内嵌 itemData。
- **Word 标记透明化**：`<w:proofErr>` / `<w:lastRenderedPageBreak>` 不再切断引文组。
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

测试：`python3 -m unittest discover -s tests`（v3 起 84 项）

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

或直接导出环境变量。**不要把 `config.local.env` 提交到公开仓库**（个人版 zip 里已含真实值）。

### Zotero MCP

本工具通过本机 Zotero MCP 服务（Streamable HTTP，默认 `127.0.0.1:23120/mcp`）写入条目。
安装方式见上方[前置依赖](#前置依赖)。请确保 Zotero 正在运行且已启用 MCP 连接器，工作流开始时会自动探测，未就绪会提示。

## 快速上手

```bash
# 1) 扫描手稿全貌（零副作用）
python3 docx_audit.py scan 论文.docx

# 2) 转换（正文；表格用 body_para_range=(0,0) + include_tables=True 单独跑）
python3 -c "
import zotero_field_insert as zfi, json
m = json.load(open('item_mapping.json'))
print(zfi.insert_zotero_fields('论文.docx', '论文_zotero.docx', (0, 165), [], m,
                               style_id='http://www.zotero.org/styles/ieee'))"

# 3) 交付前审计
python3 -c "
import zotero_field_insert as zfi, docx_audit as da, reference_ingest as ri, json
m = json.load(open('item_mapping.json'))
ok, d = zfi.verify('论文_zotero.docx', {v['itemKey'] for v in m.values()},
                   '论文.docx', m, deep=True); print(ok, d)"
python3 docx_audit.py audit 论文.docx 论文_zotero.docx item_mapping.json
```

## 工作流

1. **扫描**：`docx_audit.scan_document()`（v3 新增，见上）。
2. **前置 MCP 检查**：`check_zotero_mcp()` 探测 MCP 是否就绪；未就绪则提示用户配置。
3. **建库索引**：`build_library_doi_index()` / `doc_item_dois()`，失败条目进 `failures` 上报。
4. **解析参考文献（分支 B）**：`extract_reference_paragraphs()` → `resolve_metadata()` →
   `dedup_check()`。
5. **强制检查点**：展示汇总表 + 询问目标文件夹（现有 collection 或新建），以及每篇重复的处置。
6. **导入**：`import_to_zotero()` 写入 Zotero 到指定文件夹（幂等；导入后
   `audit_imported_by_doi()` 审计重复）。
7. **构建 mapping**：`build_item_mapping()` 产出 `item_mapping.json`；复用库内条目时
   `apply_library_data()` 以库内元数据重建 itemData。
8. **插入域代码**：`zotero_field_insert.insert_zotero_fields()`。
9. **交付前 7 项检查**（见 SKILL.md）：残留占位、逐组对账、deep 校验、死链、域计数、
   库侧处置、文件可打开 + 备份在位。

```python
from reference_ingest import (extract_reference_paragraphs, resolve_metadata, dedup_check,
                              import_to_zotero, build_item_mapping, apply_library_data,
                              check_zotero_mcp, build_library_doi_index,
                              verify_item_keys, doc_item_dois, normalize_doi)
from zotero_field_insert import insert_zotero_fields, verify
import docx_audit as da

check_zotero_mcp()
idx, failures, dups = build_library_doi_index(["COLLECTION_KEY"])
doc_idx, _, _ = doc_item_dois("论文.docx")        # 既有域引用的条目（半刷新文档取并集）

refs = extract_reference_paragraphs("论文.docx")
resolved = resolve_metadata(refs)                 # resolution_summary.json
dup = dedup_check(resolved)                       # dedup_report.json
# —— 此处为强制检查点：向用户展示 + 确认目标文件夹/去重处置 ——
plan = {k: {"action": "import", "csl": v["csl"]} for k, v in resolved.items() if v["csl"]}
imported = import_to_zotero(plan, "COLLECTION_KEY")
mapping = build_item_mapping(imported)
mapping, doi_mismatch = apply_library_data(mapping)
insert_zotero_fields("论文.docx", "论文_zotero.docx", (0, 165), [], mapping)
print(verify("论文_zotero.docx", [m["itemKey"] for m in mapping.values()], deep=True))
da.audit_output("论文.docx", "论文_zotero.docx", mapping)
ok, missing = verify_item_keys(da.zfi.document_item_keys(da.load_root("论文_zotero.docx")))
```

## 校验

`verify(out, valid_keys, src_docx, item_mapping, deep=True)` 检查：

- 结构：fldChar begin/separate/end 配对平衡、所有域的 JSON 合法（域级拼接 instrText 后解析）
- 链接：所有 URI 指向真实库条目、itemData 含 type/title/author/issued
- 零遗漏：与原文比对，每处应转换的引文都出现在输出域的显示文本里，且逐字一致
- `deep=True` 追加：citationID 唯一性、域外残留纯文字占位（跨段深度感知）

`docx_audit.audit_output(src, out, item_mapping)` 做逐组对账：显示文本、DOI 集合、URI 指向。

## 目录

```
├── SKILL.md                 # skill 触发 + 端到端工作流 + 检查点 + 避坑要点
├── docx_audit.py            # v3：扫描 / 占位提取 / 表格列审计 / 逐组对账 / 重编号 / 重链接
├── reference_ingest.py      # 前置引用摄取（解析/核实/查重/导入/索引/死链校验）
├── zotero_field_insert.py   # 域代码插入器（引文匹配 + 域构造 + 校验）
├── zotero_mcp.py            # Zotero MCP 客户端（Streamable HTTP）
├── config.py                # 配置（env + 自动推导，无硬编码 ID）
└── tests/                   # 单元测试（v3: 84 项）
```

## 版本历史

- **v3**：扫描优先工作流；citationID 自动避让 + 历史文档重编号；老式括号 DOI 全路径支持；
  跨段域深度感知；库索引/死链审计不静默；逐组对账；表格列审计；交付前 7 项检查。
- **v2**：混合文档 `skip_existing`；同段混合保护；EndNote 域剥离；表格整列 DOI；
  `apply_library_data` 以库为准。
- **v1**：基础域代码生成、纯文字参考文献解析、CrossRef/PubMed 摄取、导入检查点。

## 许可

MIT。详见 [LICENSE](LICENSE)。
