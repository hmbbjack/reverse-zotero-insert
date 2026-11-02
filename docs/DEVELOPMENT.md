# 开发与发布

[返回项目首页](../README.md) · [English](DEVELOPMENT.en.md)

## 开发

Python 3.10+，建议使用虚拟环境。仓库根目录运行：

```bash
python -m pip install ".[dev]"
python -m unittest discover -s tests -v
ruff check .
ruff format --check .
```

原有平铺模块用于兼容源码用户；wheel 把这些模块安装在 `reverse_zotero_insert` 命名空间下，避免把 `config` 等通用名称装入顶层。

`doi_utils.py` 统一 DOI 规则，`docx_io.py` 负责关闭 ZIP 文件及原子写出，`zotero_mcp.MCPClient` 管理一次顺序会话。新增业务逻辑应复用这些入口。

## 本地打包

```bash
python -m build
python scripts/build_zip.py
```

`dist/` 包含 wheel、源码包和 ZIP。ZIP 采用文件白名单，仅包含项目代码、测试、双语文档与主视觉、配置模板。
禁止把 `config.local.env`、`.env`、文稿、库数据或依赖缓存加入发布物。

安装测试须在仓库目录外运行，避免当前目录掩盖安装错误：

```bash
python -m venv /tmp/reverse-zotero-package-test
/tmp/reverse-zotero-package-test/bin/python -m pip install dist/*.whl
cd /tmp
/tmp/reverse-zotero-package-test/bin/python -m reverse_zotero_insert --help
```

Windows 虚拟环境中的 Python 位于 `Scripts/python.exe`。CI 已配置 Linux/Windows 与 Python 3.10、3.12、3.13 的测试矩阵；本地通过不代表这些远程任务已经执行。

## GitHub Release

1. 更新 `pyproject.toml`、`__init__.py` 及 MCP 客户端版本号，更新 README 版本说明。
2. 测试通过并将代码合入默认分支后，创建并推送对应版本标签，例如 `v3.1.0`。
3. `Build release` 工作流检查标签与包版本一致，执行测试并生成发布物、SHA256 校验文件。
4. 工作流创建带附件的 **Draft Release**。在 GitHub Releases 检查附件后点击 **Publish release**；此后访客才能下载版本附件。

工作流使用 GitHub 自动提供的 `GITHUB_TOKEN`，不需要把个人 token 写入仓库。
README 顶部 ZIP 按钮下载 `main` 当前源码；Release 附件是对应标签的固定版本。
当前没有自动发布到 PyPI，`pip install` 使用 GitHub URL、本地源码或 wheel。

## 已知范围

- 离线测试覆盖 DOCX XML、模拟 MCP 协议响应及主要回归场景，不替代真实 Zotero 插件和 Word Refresh 联调。
- 转换主体处理 `word/document.xml` 的正文/表格；页眉、页脚、脚注、尾注和修订标记中的引用需单独处理。
- 超链接、书签、复杂对象等结构边界可能阻止跨边界引用的自动匹配；应检查 `verify()` 报告。不要以破坏结构为代价强行转换。
- `audit_output()` 针对纯 DOI 括号组，需与 `verify(..., deep=True)` 和库内 key 校验一起使用。
- 网络写入不能整体回滚；导入中断后先核对已创建条目，再使用查重重试。
- 库索引的 collection 查询当前单次请求上限 500 条；大文件夹应显式提供完整的 `item_keys`，不能据此认定已经扫描全库。
