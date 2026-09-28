---
name: 逆向zotero插入
description: 当用户要求把参考文献以 Zotero 域代码格式插入/写入到 .docx 文件中（例如"插入 Zotero 格式文献""把引用转成 Zotero 域代码""让 Zotero 识别文档里的参考文献""用域代码替代纯文字参考文献"），或要求从纯文字的编号参考文献出发建库（例如"把参考文献插入 Zotero""从纯文本文献库入库""解析参考文献生成 Zotero 条目""给这份 docx 里的引用建 Zotero 库""resolve references and import to Zotero"），且无法直接操作 Word 的 Zotero 插件时，调用本 skill。端到端流程：扫描手稿全貌 → 联网解析纯文字编号参考文献的元数据 → 导入 Zotero（含强制用户检查点）→ 逆向 Zotero 域代码格式，直接在 docx 的 XML 中写入可被 Zotero 识别的引文域 / 参考文献表域 / 文档首选项域 → 逐组对账与交付前审计。
---

# 逆向 Zotero 插入（Reverse Zotero Field-Code Insertion）· v3

## 适用场景
把一份 .docx 里的纯文字参考文献（正文引文占位 + 文末参考文献表）转换成 Zotero 能直接识别和管理的域代码（field code），但当前环境**无法驱动 Word 的 Zotero 插件 GUI**。本 skill 通过逆向 Zotero 的域代码格式直接编辑 docx 的 XML，效果与用 Zotero 插件插入一致：在 Word 里用 Zotero 打开时，引文和文献表都是"活的"、可刷新、已链接到库条目。

## v3 新增（相对 v2）

| 能力 | 说明 |
|---|---|
| **第 0 步强制扫描** | `docx_audit.scan_document()` 一次调用拿到域清单/样式/占位组/混合段/表格引文列/citationID 分布/死链基数。**没有扫描就不要动手转换** |
| **citationID 自动避让** | 插入时自动从文档既有编号最大值 +1 起步，不再与 `citN` 世代重叠（v1/v2 恒撞号） |
| **老式 DOI 支持** | `10.1016/S0306-4530(98)00014-6` 这类含括号 DOI 在"括号组"与"整格"两条路径都能正确识别 |
| **跨段域深度感知** | 文献表域（BIBL）可跨上百段落；域外文本与残留扫描不再把文献表误判成域外占位 |
| **库索引不再静默跳错** | `build_library_doi_index()` 返回 `(index, failures, dups)`；`get_item_details` 失败必须上报 |
| **死链审计与修复** | `verify_item_keys()` 发现指向已删除条目的既有域；`relink_item_keys()` 一键重链接 |
| **逐组对账** | `audit_output()` 从原稿重新推导占位组，与输出域逐组核对（显示文本/DOI 集合/URI 指向），比数域个数强得多 |
| **安全收紧** | 分组里"形似 DOI 却查不到"的分片不再被静默截断，整组放弃（避免证据凭空消失） |
| **交付前检查清单** | 见文末"交付前硬性检查"，7 项全过才算完成 |

---

## 第 0 步：扫描（强制，不可跳过）

```bash
python3 docx_audit.py scan <file.docx>          # 人读
python3 docx_audit.py scan <file.docx> --json   # 机读
```

必看四项，它们决定后面所有分支：

1. **占位形态** —— `占位组 / 散落 DOI / 唯一 DOI`。`(doi:10.x)`、`(10.x; 10.y)`、`(FDA label, …)` 走占位通道；`(Smith, 2007)` 走作者-年份通道；两者可并存（`build_combined_finder` 会同时启用）。
2. **混合段** —— `混合段(既有域+占位)` 非空时，说明同一段里既有域又有纯文字引文。生成器的 `_process_para` 原生支持（按域区间保护），**不要**自己重写段落。
3. **表格列审计** —— 逐列给出"域 n/m、DOI 格、编号格、数据格"。**"纯数字格"是数据不是引文**（年龄 6–8、周期 14 天），只有 DOI 格与编号格要转。
4. **citationID 与 key 基数** —— `重复 >0` 说明文档已有编号重叠（历史转换造成）；`文档引用的 key 数` 用于后面的死链审计。

> 扫描零副作用，不要为了"先看一眼"就解压改写文档。

## 端到端工作流

### 分支 A：手稿用 DOI / 编号占位（可能已有部分域代码）

1. **扫描**（第 0 步）。
2. **备份**：`<name>_备份_<YYYYmmdd_HHMMSS>.docx`，与原稿同目录。转换只写新文件 `*_zotero.docx`。
3. **前置 MCP 检查**：`reference_ingest.check_zotero_mcp()`。未就绪则停下提醒用户启动 Zotero。
4. **建库索引**：`build_library_doi_index(collection_keys=[...], csl_builder=det_to_csl)` → 拿 `(index, failures, dups)`。**`failures` 必须逐条上报用户**（多为你已删除的条目）；`dups` 同 DOI 多条目要问用户合并到哪条。索引范围要**全库扫**：占位 DOI 常落在你没料到的文件夹里。
5. **半刷新文档取并集**：既有域引用的条目也要进 mapping —— `doc_item_dois(docx)`。只取占位集合会让"仅被既有域引用"的文献看起来像映射外。
6. **补齐缺口**：未命中的 DOI 走 CrossRef/PubMed 解析（`search_crossref_by_doi`）。老式括号 DOI 要用 v3 的 `normalize_doi()` 归一后再查。
7. **强制检查点（硬停）**：展示解析汇总 + 目标文件夹（列出现有 collection + "新建文件夹"）+ 每条去重处置。**未获明确选择不得 `write_item` / `add_items_to_collection`。**
8. **导入**：`import_to_zotero(plan, collection_key)`（幂等，`precheck=True`）。导入后 `audit_imported_by_doi()` 审计重复，命中 >1 的列出让用户手动删（MCP 无删除工具）。
9. **以库为准**：`apply_library_data(item_mapping)` 用库内元数据重建 itemData，人工核对返回的 DOI 不一致清单。
10. **干跑**：`finder` 在目标段落范围上跑一遍，确认命中数 == 预期组数、且范围外零命中（防止作者-年份通道误伤）。表格用 `body_para_range=(0,0)` + `include_tables=True` 单独跑。
11. **插入**：`insert_zotero_fields(...)`（正文）/ 见下（表格）。`ref_para_indices=[]`（文献表域已存在时不重复添加）。
12. **交付前检查**（见文末清单）。

### 分支 B：纯文字编号参考文献（需先建库）

在分支 A 之前插入：解析编号块（`extract_reference_paragraphs`）→ `resolve_metadata`（CrossRef/PubMed 搜索 API）→ `dedup_check` → **强制检查点** → `import_to_zotero` → `build_item_mapping` → 回到分支 A 第 9 步。

> **混合文档支持**：`insert_zotero_fields` 默认 `skip_existing=True`，已含引文域的段落原样保留，bibliography/prefs 域已存在则不重复添加。`detect_zotero_fields()` 可单独预览。

### 表格的处理

- **引文列整格即 DOI**：`insert_table_citations(src, out, mapping, doi_column, refnum_column)`。
- **同列混合（部分格已是域、部分是纯文字占位）**：走通用路径 `insert_zotero_fields(..., body_para_range=(0,0), include_tables=True)`，`_process_para` 按域区间保护既有格。
- **表格单独成文件**（引文列与编号列跨文件对应）：两个文件**各自跑一次插入**，编号由 Zotero 刷新时按各文档内引文顺序重排；两文件编号方案保持一致（都用 IEEE 或都用 Vancouver），否则刷新后对不上。

### 引用格式

沿用文档自带 PREF 域的样式（扫描结果里的 `style`）。文档没有 PREF 域时：先问用户目标期刊/通用格式，`resolve_style_id()` 命中别名则得 styleID，未命中则联网搜该期刊的 Zotero CSL 样式。表格文件与主稿样式不一致时**必须问用户**。

### 强制检查点细则

Claude 用 AskUserQuestion 式交互（编号菜单）询问，等用户回答后再继续。目标文件夹选项 = 每个现有 collection（label=name，value=key）+ 末尾"新建文件夹"哨兵（选后再问名字，调 MCP `create_collection` 拿 key）。对 `DUP/PROB-DUP` 条目问 `(a)跳过 (b)导入为新条目 (c)复用库中已有条目`；对 `LOW` 置信度先联网复核修正再确认。用户确认后把选择映射进 `import_plan[ref_key]["action"]` ∈ `skip|import|reuse`。

## 核心原理（域代码格式）

三类域，存放在 `<w:instrText>` 里，每域由 begin / instrText / separate / 显示文本 / end 五个 run 组成：

1. **引文域** `ADDIN ZOTERO_ITEM CSL_CITATION {json}` —— `{"citationID","properties":{"formattedCitation","plainCitation","dontUpdate","noteIndex"},"citationItems":[{"id","uris","itemData"}],"schema"}`。`dontUpdate:false` 让引文随样式重排；显示文本缓存保留占位原文，Refresh 后自动变编号。
2. **参考文献表域** `ADDIN ZOTERO_BIBL {json} CSL_BIBLIOGRAPHY` —— **尾部 `CSL_BIBLIOGRAPHY` 不能漏**。
3. **文档首选项域** `ADDIN ZOTERO_DOCUMENT_PREFERENCES {json}`，书签 `ZOTERO_PREF`；文献表域用书签 `ZOTERO_BREF`。

URI 形如 `http://zotero.org/users/<userID>/items/<itemKey>`。`userID` 与 API 端点由 `config.py` 从环境变量 / `config.local.env` / `~/Zotero/zotero.sqlite` 读取，代码零硬编码。

## ⚠️ 避坑要点（每条都对应一次真实事故）

**域与扫描**
- **域深度必须跨段落连续传递**。文献表域跨成百上千段；逐段把深度重置为 0，会把文献表正文里的 `https://doi.org/…` 判成"域外残留占位"，虚增上百条假命中。用 `zfi.outside_field_paragraphs()` / `zfi.collect_fields()`，不要自己重写。
- **域内缓存显示文本含 DOI 字样**。转换时显示文本 = 占位原文，所以"已转换"的引文在纯文本扫描里仍长得像占位。判断"是否还有真占位"只能用域深度感知扫描。
- **一个域的指令可能被 Word 拆成多个 `<w:instrText>`**。必须域级拼接后再 `json.loads`；逐元素解析会把长 JSON 误报成截断错误。
- **`collect_fields` 的 instrText 存元素不是字符串**（要就地重写就存元素）。
- **Word 的 `<w:proofErr>` / `<w:lastRenderedPageBreak>` 会切断引文组**：转换器已作透明节点丢弃，别改回去。
- **`w:br` 软换行不产生字符**：多 DOI 单元格拼接文本会粘连成 `…012doi:10.x`，匹配失败。提取单元格文本时把 `br` 记作 `\n`。
- **按段落号配对原稿与输出是不可靠的**：insert 会在文首插入首选项域段落，输出与原稿的段落序号整体位移。逐组对账按**显示文本多重集**配对。

**citationID**
- **新域编号必须避开既有编号**。v3 已自动从 max+1 起步；历史文档（v1/v2 插过的）用 `docx_audit.renumber_citation_ids()` 修。
- **跨版本复制段落会造成编号世代重叠**：同一篇稿子改到 v3 再转一次，旧域还带着 v1 的 `cit0..citN`。交付前 `verify(deep=True)` 或 `scan_document()` 看"重复"。
- **插入返回的 `cit_n` 是"新增域数"**，不是编号终点；编号起点是内部 `cit_start`。

**DOI 解析**
- **老式 Elsevier DOI 内含括号**：`10.1016/S0306-4530(98)00014-6`。朴素的 `10\.\d{4,9}/[^\s,;)]+` 会截断成 `10.1016/s0306-4530`，查库必落空；朴素的 `\(([^()]*)\)` 也匹配不到整组。一律用 `normalize_doi()` 与 `docx_audit.extract_doi_groups()`。
- **DOI 归一化再比对**：`https://doi.org/…`、`doi:…`、`(10.x)`、尾部 `.`/`;` 都要归一，否则同一篇文献查不中、重复导入。
- **CrossRef 标题含 HTML 标记**（`<i>…</i>`）：`crossref_to_csl` 已剥除，自行解析时要清理。
- **死链 DOI** 多为笔误；`resolve_error` 要显式标记，结合上下文用 CrossRef 搜索找正确 DOI，并向用户确认后同步改正文文字。

**库与查重**
- **`search_library` 结果不含 DOI 字段**（minimal/standard 都不含），必须 `get_item_details(mode=complete)` 逐条确认。
- **MCP 返回形态不统一**：`get_collection_items` 返回数组，`search_library` 返回 `{"results": …}`。统一用 `reference_ingest._results()` 解析；解析不了会**静默建出空索引**。
- **建索引遇到 `get_item_details` 报错不许静默 `continue`**：那正是"用户已删除的条目"，必须进 `failures` 上报。
- **同 DOI 多条目**：按相关度排序的搜索结果会掩盖第二个重复；标题大小写/Unicode 连字符差异会让标题搜索漏掉库内条目，标题未 exact 命中时**总是**再走 DOI 路径。
- **MCP 没有 `delete_item`**：去重只能 skip / reuse / import-as-new，绝不删除既有条目。
- **`write_item` 参数位置**：`itemType` 必须是顶层参数；`creators` 独立列表且空列表必须整个省略。
- **半刷新手稿的 mapping 取并集**（占位 DOI ∪ 既有域引用）。

**作者-年份通道**
- 不要用 `Cohen`、`e.g.`、`i.e.` 等关键字过滤"非引文"括号组（会误杀 Baron-Cohen、"e.g., Smith, 2007"），靠"作者姓+年份"验证排除。
- 多词姓氏（De Vries、van Rijn）要同时尝试全姓与末词。
- 引文匹配不要用于表格的数据列：数字格会撞上"纯数字"模式，先用 `table_column_audit()` 分清列。

## 交付前硬性检查（7 项）

1. `docx_audit.scan_document(out)` → 占位组 0、散落 DOI 0。
2. `docx_audit.audit_output(src, out, item_mapping)` → `problems == []`（逐组对账）。
3. `zfi.verify(out, valid_keys, src, item_mapping, deep=True)` → `ok == True`（含 citationID 撞号、域外残留、坏 JSON、坏 URI）。
4. `reference_ingest.verify_item_keys(文档全部 key)` → `missing == []`；有 missing 用 `relink_item_keys()` 修（库里找同 DOI 的规范条目重链接），修不了才向用户报告。
5. 域计数符合预期：`域 == 原域 + 新域`，fldChar begin/end 平衡。
6. 库侧：索引的 `failures` 已上报、`dups` 已处置、`audit_imported_by_doi` 无 >1 命中。
7. python-docx 能打开输出文件；备份文件在位。

## 脚本

| 文件 | 作用 |
|---|---|
| `zotero_field_insert.py` | 域代码生成器：匹配器、插入（正文/表格）、`verify()` |
| `reference_ingest.py` | 前端摄取：解析、联网核实、建库、导入、mapping 构建、库索引、死链校验 |
| `docx_audit.py` | **v3**：扫描、占位提取、表格列审计、逐组对账、重编号、重链接、CLI |
| `zotero_mcp.py` | Zotero MCP 的 JSON-RPC 客户端 |
| `config.py` | 个人标识（userID / API 端点 / 限流）读取，代码零硬编码 |
| `tests/` | `python3 -m unittest discover -s tests`（v3 起 84 项） |

## 典型调用

```python
import sys; sys.path.insert(0, '<skill目录>')
import docx_audit as da, zotero_field_insert as zfi, reference_ingest as ri

# 第 0 步：扫描
da.scan_document(src, verbose=True)

# 库索引（失败必须上报）
idx, failures, dups = ri.build_library_doi_index(['COLLECTION'], csl_builder=det_to_csl)
doc_idx, dfail, ddup = ri.doc_item_dois(src, csl_builder=det_to_csl)   # 既有域引用的条目
mapping = {f'ref_{d}': {'itemKey': e['itemKey'], 'uri': e['uri'], 'doi': e['doi'],
                         'csl': e['csl'], 'first_author': ..., 'year': ...}
           for d, e in {**idx, **doc_idx}.items()}

# 插入（正文；citationID 自动避让既有编号）
cit_n, bib_n = zfi.insert_zotero_fields(
    src_docx=src, out_docx=out, body_para_range=(0, body_end), ref_para_indices=[],
    item_mapping=mapping, style_id=style_id, style_locale='en-US',
    skip_existing=True, include_tables=False)

# 插入（仅表格）
zfi.insert_zotero_fields(src_docx=tbl, out_docx=tbl_out, body_para_range=(0, 0),
                         ref_para_indices=[], item_mapping=mapping, style_id=style_id,
                         include_tables=True)

# 交付前
ok, detail = zfi.verify(out, valid_keys, src, mapping, deep=True)
da.audit_output(src, out, mapping)
ok_keys, missing = ri.verify_item_keys(zfi.document_item_keys(da.load_root(out)))
```

## 版本历史

- **v3**（2026-09）：扫描优先工作流；citationID 自动避让与历史文档重编号工具；老式括号 DOI 全路径支持；跨段域深度感知；库索引/死链审计不静默；逐组对账；表格列审计；交付前 7 项检查；84 项测试。
- **v2**：混合文档 `skip_existing`；同段混合保护；EndNote 域剥离；表格整列 DOI；`apply_library_data` 以库为准。
- **v1**：基础域代码生成、纯文字参考文献解析、CrossRef/PubMed 摄取、导入检查点。
