# README 视觉与文档设计

默认首页为中文 `README.md`；英文版为 `README.en.md`，两者在顶部互相链接。

版式参考了以下项目的公开 README，但没有复用它们的标志、插图或宣传数据：

- [uv](https://github.com/astral-sh/uv)：开头直接说明用途，随后给出安装入口。
- [Rich](https://github.com/Textualize/rich)：先用视觉展示能力，再展开功能说明。
- [FastAPI](https://github.com/fastapi/fastapi)：居中的品牌区、少量徽章和清楚的文档导航。

本项目使用暖纸色、墨蓝与钴蓝，配合手稿、引用连接线和条目卡片，表达“把纯文本引用连接到结构化文献记录”的主题。中文和英文横幅分别生成，保持同一构图。

- `assets/readme-hero-zh.png`：中文版主视觉。
- `assets/readme-hero-en.png`：英文版主视觉。
- 横幅由 AI 生成，是概念插图，不是 Word/Zotero 界面截图。
- 精确的业务流程用两版 README 中的 Mermaid 图表达。
- 首页保留用途、安装、工作流、配置与能力范围；完整示例和维护说明放在 `docs/`。
- 不使用未经验证的下载量、覆盖率或“CI 已通过”徽章。

更新安装方式、功能边界或发布状态时，需同步两版 README 与相应语言的使用指南。
