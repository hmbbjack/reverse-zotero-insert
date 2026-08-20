---
name: 逆向zotero插入
description: 当用户要求把参考文献以 Zotero 域代码格式插入/写入到 .docx 文件中（例如"插入 Zotero 格式文献""把引用转成 Zotero 域代码""让 Zotero 识别文档里的参考文献""用域代码替代纯文字参考文献"），或要求从纯文字的编号参考文献出发建库（例如"把参考文献插入 Zotero""从纯文本文献库入库""解析参考文献生成 Zotero 条目""给这份 docx 里的引用建 Zotero 库""resolve references and import to Zotero"），且无法直接操作 Word 的 Zotero 插件时，调用本 skill。端到端流程：前置 MCP 配置检查 → 联网解析纯文字编号参考文献的元数据 → 导入 Zotero（含强制用户检查点）→ 逆向 Zotero 域代码格式，直接在 docx 的 XML 中写入可被 Zotero 识别的引文域 / 参考文献表域 / 文档首选项域。
---

# 逆向 Zotero 插入（Reverse Zotero Field-Code Insertion）

## 适用场景
用户希望把一份 .docx 里的纯文字参考文献（正文引文 + 文末参考文献表）转换成 Zotero 能直接识别和管理的域代码（field code），但当前环境**无法驱动 Word 的 Zotero 插件 GUI**（例如在命令行/自动化场景）。本 skill 通过逆向 Zotero 的域代码格式，直接编辑 docx 的 XML 写入域代码，效果与用 Zotero 插件插入一致——在 Word 里用 Zotero 打开时，引文和参考文献表都是"活的"、可刷新、已链接到 Zotero 库条目。

## 端到端工作流（含前端引用摄取）

当输入是**纯文字编号参考文献**（引用无 DOI、需先建库）时，走完整端到端流程；若条目已在 Zotero 库中，可直接跳到第 7 步。

1. **前置 Zotero MCP 配置检查**：先 `reference_ingest.check_zotero_mcp()` 探测 MCP 是否就绪。若 `ok=False`，**提醒用户"Zotero MCP 未配置好"**——需确认 Zotero 正在运行且已启用 MCP 连接器（监听 `127.0.0.1:23120`），或设置 `ZOTERO_MCP_URL` 环境变量；**未就绪则不进入摄取流程**，先等用户处理。
2. **解析编号参考文献**：`reference_ingest.extract_reference_paragraphs()` 读出编号参考文献块（`[n]`/`(n)`/`n.`/`n)` 引导）。
3. **批量解析元数据**：`reference_ingest.resolve_metadata()` 用 CrossRef/PubMed **搜索 API**（非 DOI 反查）逐条核实，产出 `resolution_summary.json`（含置信度 HIGH/UNKNOWN/LOW）。
4. **去重检查**：`reference_ingest.dedup_check()` 用 MCP `search_library` 按标题/DOI 查重，产出 `dedup_report.json`（NEW/DUP/PROB-DUP）。
5. **强制用户检查点（硬停）**：展示元数据汇总表，并**询问目标文件夹**（列出现有 collection + "新建文件夹"选项）与每篇重复的处置；低置信度条目先 `WebSearch` 复核。**未获用户对「目标文件夹 + 每篇去重处置 + 低置信度确认」的明确选择，不得调用 `write_item`/`add_items_to_collection`/`import_to_zotero`。**
6. **导入 Zotero**：`reference_ingest.import_to_zotero(import_plan, collection_key)` 写入到用户选定的文件夹（幂等：`precheck=True` 先查重，已有自动转复用）。导入后 `audit_imported_by_doi(resolved)` 审计重复（命中 >1 的条目列出让用户手动删，MCP 无删除工具）。
7. **构建 mapping**：`reference_ingest.build_item_mapping()` 产出 `item_mapping.json`（格式与旧流程逐字节兼容）。复用库内条目时建议 `apply_library_data(item_mapping)` 以库内元数据重建 itemData（CrossRef 与库常有出入），并人工核对返回的 DOI 不一致清单。
8. **强制检查点：引用格式/期刊选择**：向用户询问目标**期刊或通用引用格式**（如 Nature、IEEE、Vancouver、GB/T 7714、APA…）。`reference_ingest.resolve_style_id()` 命中内置别名则直接得 styleID；未命中则 Claude **联网搜索**该期刊对应的 Zotero CSL 样式，得到 `style_id`（及 locale）。**未获用户对格式的选择，不得插入域代码。**
9. **插入域代码 + 校验**：`zotero_field_insert.insert_zotero_fields(..., style_id, style_locale)` + `verify()`。

> **混合文档支持**：若文档里**既有已有 Zotero 域代码、又有纯文字参考文献**（常见于插件插入过部分引文），`insert_zotero_fields` 默认 `skip_existing=True`，会自动检测并**原样保留**已含引文域的段落，只转换纯文字部分；bibliography/prefs 域已存在则不重复添加。`detect_zotero_fields()` 可单独预览检测结果。

> **强制检查点细则**：Claude 在第 5 步停下，用 AskUserQuestion 式交互（编号菜单）询问，等用户回答后再继续。目标文件夹选项 = 每个现有 collection（label=name，value=key）+ 末尾"新建文件夹"哨兵（选后 Claude 再问名字，调 MCP `create_collection` 拿 key）。对 `DUP/PROB-DUP` 条目问 `(a)跳过 (b)导入为新条目 (c)复用库中已有条目`；对 `LOW` 置信度先 `WebSearch` 复核修正再确认。用户确认后，把选择映射进 `import_plan[ref_key]["action"]` ∈ `skip|import|reuse`（reuse 带 `zotero_key`）。第 8 步的格式选择同样等用户回答后再插入。

## 核心原理（逆向得到的格式）

Zotero 在 Word 文档（`word/document.xml`）中用三类域代码，存放在 `<w:instrText>` 里。每个域由一组 run 组成：
`<w:r><w:fldChar w:fldCharType="begin"/></w:r>` → `<w:r><w:instrText xml:space="preserve"> ADDIN … </w:instrText></w:r>` → `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` → `<w:r><w:t xml:space="preserve">显示文本</w:t></w:r>` → `<w:r><w:fldChar w:fldCharType="end"/></w:r>`

1. **引文域**：`ADDIN ZOTERO_ITEM CSL_CITATION {json}`
   - json = `{"citationID":"…","properties":{"formattedCitation":"显示文本","plainCitation":"显示文本","dontUpdate":false,"noteIndex":0},"citationItems":[{"id":N,"uris":["http://zotero.org/users/<userID>/items/<itemKey>"],"itemData":<CSL-JSON>}],"schema":"https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}`
   - 默认 `dontUpdate:false`（引文随 Zotero 样式自动重排）；仅当要**冻结**某条引文外观不改时才设 `true`。
2. **参考文献表域**：`ADDIN ZOTERO_BIBL {json} CSL_BIBLIOGRAPHY`（**注意尾部 `CSL_BIBLIOGRAPHY`**，不能漏）
   - json = `{"uncited":[["uri"],…],"omitted":[],"custom":[]}`。参考文献表内容由引文域 + uncited 重新生成；域的显示文本只是缓存。
3. **文档首选项域**：`ADDIN ZOTERO_DOCUMENT_PREFERENCES {json}`，用书签 `ZOTERO_PREF` 包裹；参考文献表域用书签 `ZOTERO_BREF` 包裹。
   - json = `{"style":{"styleID":"http://www.zotero.org/styles/apa","locale":"en-US","hasBibliography":true,"bibliographyStyleHasBeenSet":false},"prefs":{"fieldType":"Field","automaticJournalAbbreviations":false},"sessionID":"…","zoteroVersion":"9.0.6","dataVersion":4}`

> 格式来源：从 `/Applications/Zotero.app/…/app/omni.ja` 里的 `xpcom/integration.js`（`setCode`、`Citation.toJSON`、`Bibliography.serialize`）和 `libZoteroWordIntegration.dylib` 的字符串（`ITEM CSL_CITATION ` / `BIBL ` / `DOCUMENT_PREFERENCES ` / ` ADDIN ZOTERO_`）逆向确认。

## 关键参数获取

- **userID（URI 前缀）**：Zotero 个人库条目的 URI = `http://zotero.org/users/<userID>/items/<itemKey>`。`<userID>` 由 `config.py` 从环境变量 `ZOTERO_USER_ID` 或 Zotero 数据库 `users` 表读取（DB 在 `~/Zotero/zotero.sqlite`，Zotero 运行时被锁——复制一份快照再读）。已同步库用真实数字 userID；未同步库用 `users/local/<localKey>`。
- **itemKey**：通过 Zotero MCP（`http://127.0.0.1:23120/mcp`，JSON-RPC）的 `write_item`(action=create) 或 `search_library`+`get_item_details`(mode=complete 才返回 DOI) 获取。
- **CSL-JSON itemData**：用 CrossRef 搜索 API（`https://api.crossref.org/works?query.bibliographic=...`）按书目信息解析得到完整元数据（含完整作者列表，非"et al."）；无 DOI 的用 PubMed 兜底。
- **配置**：所有个人标识（userID、URI 前缀、CrossRef mailto、NCBI key、MCP URL）由 `config.py` 统一读取，按 `KEY=VALUE` 写入环境变量或 `config.local.env`（放代码同目录即可被自动加载），代码零硬编码。未配置时自动从 `~/Zotero/zotero.sqlite` 推导 userID。

## 工作流程

1. **解析文档**：用 python-docx 读出每个段落的文字、定位正文引文位置和参考文献表段落范围。
2. **建立引文匹配器**：对每篇文档，按其参考文献集构建 `{年份: [(第一作者姓, itemKey)]}`。匹配时**大小写不敏感**（CrossRef 偶尔返回全大写姓，如 PHOENIX），且匹配"姓 + 边界"（后面跟 et al / and / 逗号-位于段首）。
3. **生成域代码**：把每处引文的文字替换为引文域（显示文本=原文，默认 `dontUpdate:false` 让引文随样式自动重排；如需冻结单条外观可设 `true`，并通过 uris 链接到库条目，itemData 内嵌作兜底）。参考文献表段落用参考文献表域包裹。文首插入首选项域（隐藏段落，书签 ZOTERO_PREF）。
4. **写回 docx**：只替换 `word/document.xml`，其余 zip 部分原样保留。
5. **校验**：`verify()` 检查 fldChar begin/separate/end 配对平衡；所有 JSON 合法；所有 URI 指向真实库条目；itemData 含 type/title/author/issued。若传入 `src_docx` 与 `item_mapping`,还会比对原文做**引文零遗漏 + 显示文本逐字一致**校验（用引文匹配器在原文里找出所有"作者+年份",逐一核对是否都出现在输出域的显示文本中）。

## ⚠️ 避坑要点（实测踩过的坑）

- **引文匹配过滤**：不要用 `Cohen`、`e.g.`、`i.e.` 等关键字过滤"非引文"括号组——`Cohen` 会误杀"Baron-Cohen"，`e.g.` 会误杀"(e.g., Smith, 2007)"。改为靠"作者姓+年份"验证来排除非引文。
- **占位符风格的正文（DOI/注册号占位）**：手稿正文未必用"作者+年份"，可能是 `(doi:10.x; 10.y)`、`(10.x)`、`(FDA label, …)`、`(NCT…)`、`(EU CT …)`、`(公司, 年份)` 等。`find_citations` 只认作者-年份，对这类文档一无所获；用 `build_doi_matcher` + `find_doi_citations`（DOI 组）和 `citation_text_map`（精确文本括号）匹配，`insert_zotero_fields` 已内置（mapping 条目含 `doi` 即自动启用）。表格"整列即 DOI"用 `insert_table_citations(refnum_column=None)` 或 `include_tables=True`。
- **Word 的 `<w:proofErr>` 会切断引文组**：Word 随文插入拼写/渲染标记（proofErr、lastRenderedPageBreak），若按普通元素切分文本块，跨标记的括号引文组会整组漏转（实测一份文档 262 处标记、漏掉 22 组）。转换器已把它们作透明节点丢弃，勿改回。
- **`search_library` 结果不含 DOI 字段**：minimal 和 standard 模式都不返回 DOI。任何 DOI 匹配必须 `get_item_details(mode=complete)` 逐条确认，否则永远落空（曾因此漏检 4 个库内已有条目、重复导入）。
- **重新生成尽量用干净原件**：生成器默认 `skip_existing=True`，会检测并跳过已含 Zotero 域的段落/域，支持"既有域代码又有纯文字"的混合文档，可安全地在原文件上处理。但若把"已写入域代码的文件"整份再走一次全量转换，仍可能重复写入——只对纯文字部分重新生成。
- **多词姓氏**：De Vries、van Rijn、Orban de Xivry 等——匹配器要同时尝试"全姓"和"末词"。
- **dontUpdate**：默认设 `false`，使引文随 Zotero 里选择的样式自动重排（用户改样式后刷新即可生效）。只有在需要"冻结"某条引文外观、不让 Zotero 重排时才设 `true`（此时该条引文被标记为手动编辑，不会随样式变化）。注意：表格里 Ref# 列若显示"[1]"，`dontUpdate:false` 下改用 APA 样式刷新会变成"(Choi et al., 2016)"--若想保持数字编号，选用数字型样式(IEEE/Vancouver)即可。
- **域代码 JSON 转义**：写入 `<w:instrText>` 时对 `&` `<` `>` 做 XML 转义。
- **`write_item` 参数位置**：`itemType` 必须是**顶层**参数，不能放进 `fields`（放进 fields 会报错）；`creators` 是独立列表，且**空列表必须整个省略**（Zotero 报 "Creator names cannot be empty"，会让批量导入中途崩掉；导入默认 `precheck=True` 幂等，中途失败重跑不会重复建条目）。
- **无 `delete_item` 工具**：去重只能靠 skip/reuse/import-as-new，**绝不删除**已有条目；导入后用 `audit_imported_by_doi()` 审计重复（命中数 >1 只能列出让用户手动清理）。
- **去重可能多轮**：`search_library` 按相关度排序，首个 probable 命中可能掩盖第二个重复。标题大小写/Unicode 连字符差异会让标题搜索漏掉库内条目，所以标题未 exact 命中时**总是**再走 DOI 路径（详情确认）。
- **复用库内条目时以库为准**：CrossRef 数据与库内常有出入（标题大小写、连字符、缺录 DOI/卷期页）。插入前用 `apply_library_data(item_mapping)` 把 itemData 重建为库内元数据；返回的 DOI 不一致清单要人工核对（多为库内错录或 DOI 笔误）。
- **CrossRef 标题含 HTML 标记**：`<i>…</i>` 等（`crossref_to_csl` 已自动剥除）；自己解析时记得清理，否则标记会写进 Zotero 标题。
- **死链 DOI**：DOI 404 多为笔误（如 `…-020-20296-5` vs `…-020-20604-3`）。`resolve_metadata` 以 `resolve_error` 显式标记，勿静默跳过--结合上下文用 CrossRef 搜索找正确 DOI，向用户确认后替换（含正文文字）。
- **verify 与 abstract**：真实 Zotero 域经常内嵌 abstract，不算缺陷（`with_abstract` 仅计数提示）；但新生成的域仍应剔除 abstract（生成器已做）。
- **限流**：CrossRef 礼貌池 ~1/1.2s、NCBI ~1/0.4s；429 退避重试。CrossRef 用 `CROSSREF_MAILTO` 降低限流，NCBI 用 `NCBI_API_KEY`。

## 脚本

同目录 `zotero_field_insert.py` 是可复用的生成器（含匹配器、域代码生成、校验），`reference_ingest.py` 是前端摄取（解析/解析元数据/去重/导入/mapping）。用法见文件顶部注释。

**两段式典型调用**（前端摄取 → 插入）：

```python
# 阶段一：前端摄取（纯文字编号参考文献 -> Zotero 库 -> item_mapping.json）
from reference_ingest import check_zotero_mcp, extract_reference_paragraphs, \
    resolve_metadata, dedup_check, import_to_zotero, build_item_mapping, \
    audit_imported_by_doi, apply_library_data

check_zotero_mcp()                              # 前置 MCP 配置检查（未就绪则提醒用户）
refs = extract_reference_paragraphs('论文.docx') # 解析编号参考文献
resolved = resolve_metadata(refs)               # 联网核实元数据 -> resolution_summary.json
dup = dedup_check(resolved)                     # 去重检查 -> dedup_report.json
# ---- 强制检查点：展示汇总表，询问用户目标文件夹 + 每篇去重处置（未确认不得导入）----
plan = {k: {"action": "import", "csl": v["csl"]} for k, v in resolved.items() if v["csl"]}
imported, warnings = import_to_zotero(plan, 'COLLECTION_KEY')  # 用户选定的文件夹
if warnings: print(warnings)  # 若有 write_item 未取到 key 的条目，需人工处理
audit = audit_imported_by_doi(resolved)        # 导入后审计重复 -> doi_audit.json
item_mapping = build_item_mapping(imported)     # -> item_mapping.json
item_mapping, doi_mismatch = apply_library_data(item_mapping)  # 以库为准重建 itemData

# 阶段二：域代码插入 + 校验
from zotero_field_insert import insert_zotero_fields, verify
from reference_ingest import resolve_style_id
# ---- 强制检查点：询问用户目标期刊/通用格式（如 Nature、IEEE、GB/T 7714）----
journal = 'Nature'                       # 用户回答；未命中别名表时 Claude 联网搜索 styleID
style_id, style_locale = resolve_style_id(journal)
insert_zotero_fields(
    src_docx='论文.docx',               # 原始文件（混合文档会自动保留已有域代码）
    out_docx='论文_zotero.docx',
    body_para_range=(1, 118),           # 正文段落范围（含）
    ref_para_indices=list(range(119,180)),  # 参考文献表段落索引
    item_mapping=item_mapping,          # {ref_key: {itemKey, uri, csl, doi}}
    style_id=style_id,
    style_locale=style_locale,
    # 占位符风格正文（可选）：citation_text_map={'FDA label, minocycline': ['fdamino'], ...}
    # 表格"整列即 DOI"时加 include_tables=True
)
```

`item_mapping.json` 需事先用前端摄取阶段（Zotero MCP + CrossRef/PubMed）构建。若条目已在库中，可跳过阶段一，直接按旧流程用 MCP + CrossRef 手工构建 mapping 后进阶段二。对**表格类**文档（如数据提取表），把引文域放到"编号/Ref"列而不是作者列——见 `insert_table_citations` 的说明。
