# 使用指南

[返回项目首页](../README.md) · [English](USAGE.en.md) · [开发与发布](DEVELOPMENT.md)

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

使用源码时，将 `config.local.env` 放在 `config.py` 同目录；pip 安装后推荐直接设置环境变量。
**不要提交个人配置**。标准发布包只包含 `config.example.env`，不会收录本地配置或用户文稿。

### Zotero MCP

本工具通过本机 Zotero MCP 服务（Streamable HTTP，默认 `127.0.0.1:23120/mcp`）写入条目。
安装方式见[项目首页的 Zotero 配置](../README.md#zotero-setup)。请确保 Zotero 正在运行且已启用 MCP 连接器，工作流开始时会自动探测，未就绪会提示。

## 转换与校验

以下转换示例假设正文为第 0–164 段，`item_mapping.json` 已包含真实 Zotero 条目；请先扫描，再按实际文档调整范围。

```bash
# 1) 扫描手稿全貌（零副作用）
reverse-zotero-insert scan 论文.docx

# 2) 转换（正文；表格用 body_para_range=(0,0) + include_tables=True 单独跑）
python3 -c "
from reverse_zotero_insert import zotero_field_insert as zfi
import json
m = json.load(open('item_mapping.json'))
print(zfi.insert_zotero_fields('论文.docx', '论文_zotero.docx', (0, 165), [], m,
                               style_id='http://www.zotero.org/styles/ieee'))"

# 3) 交付前审计
python3 -c "
from reverse_zotero_insert import zotero_field_insert as zfi, docx_audit as da
import json
m = json.load(open('item_mapping.json'))
ok, d = zfi.verify('论文_zotero.docx', {v['itemKey'] for v in m.values()},
                   '论文.docx', m, deep=True); print(ok, d)"
reverse-zotero-insert audit 论文.docx 论文_zotero.docx item_mapping.json
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
from reverse_zotero_insert.reference_ingest import (
    extract_reference_paragraphs,
    resolve_metadata,
    dedup_check,
    import_to_zotero,
    build_item_mapping,
    apply_library_data,
    check_zotero_mcp,
    build_library_doi_index,
    verify_item_keys,
    doc_item_dois,
    normalize_doi,
)
from reverse_zotero_insert.zotero_field_insert import insert_zotero_fields, verify
from reverse_zotero_insert import docx_audit as da

status = check_zotero_mcp()
if not status["ok"]:
    raise RuntimeError(status["hint"])
idx, failures, dups = build_library_doi_index(["COLLECTION_KEY"])
doc_idx, _, _ = doc_item_dois("论文.docx")  # 既有域引用的条目（半刷新文档取并集）

refs = extract_reference_paragraphs("论文.docx")
resolved = resolve_metadata(refs)  # resolution_summary.json
dup = dedup_check(resolved)  # dedup_report.json
# —— 此处为强制检查点：向用户展示 + 确认目标文件夹/去重处置 ——
plan = {k: {"action": "import", "csl": v["csl"]} for k, v in resolved.items() if v["csl"]}
imported, import_warnings = import_to_zotero(plan, "COLLECTION_KEY")
if import_warnings:
    raise RuntimeError(import_warnings)
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
- 链接：所有 URI 的 key 属于调用者传入的有效集合、itemData 含 type/title/author/issued；真实库内存在性另用 `verify_item_keys()` 检查
- 零遗漏：与原文比对，每处应转换的引文都出现在输出域的显示文本里，且逐字一致
- `deep=True` 追加：citationID 唯一性、域外残留纯文字占位（跨段深度感知）

`docx_audit.audit_output(src, out, item_mapping)` 对纯 DOI 括号组逐组核对显示文本、DOI 集合与 URI 指向；作者年份、独立 DOI 和精确文本引用的遗漏检查由 `verify()` 完成。

命令行 `audit` 的退出码为：`0` 无异常，`1` 对账失败，`2` 参数或文件错误。
重编号与重链接支持 `--dry-run`。转换仍建议写入新文件；内部采用临时文件和原子替换，即使误用相同输入/输出路径，也不会先截断原稿。

查重请求失败会抛出异常并停止当前导入流程；`apply_library_data()` 失败时不会留下部分更新的 mapping。
网络写入不具备事务回滚能力：失败前已经成功创建的条目仍在 Zotero 中，重试前需核查。
