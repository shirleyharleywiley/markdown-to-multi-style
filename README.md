# Markdown Style Converter

一键将 Markdown 转换为带精美样式的 Word 或 HTML，支持 8 种专业版式。

**专为 AI 产出内容设计**，双语文档、中英混杂、表格等格式都能完美呈现。

## 8 种版式

| 版式 | 配色 | 适合场景 |
|------|------|----------|
| left-block | 薄荷绿 | 技术文档、公众号 |
| bottom-deco | 粉红 | 技术博客 |
| gradient-pill | 柠檬黄 | 公众号 |
| stripe-docs | 紫色 | API文档、开发指南 |
| cmd-markdown | 酒红 | 团队笔记 |
| notion | 蓝色 | 双语内容 |
| medium | 青绿 | 长文阅读 |
| linear | 浅绿 | 代码密集 |

## 快速开始

```
/markdown-to-word-multi-style /path/to/article.md
```

## 输出格式

- **Word (.docx)** — Stripe Docs / Notion（表格显示效果极佳）
- **HTML** — 8 种版式全部支持

> 示例文件在 `examples/` 目录，可直接打开查看效果

## 避坑指南

| 版式 | 避免用于 |
|------|----------|
| gradient-pill | 含表格的文档 |
| bottom-deco | 正式工作报告 |
| left-block | 标题密集的文档 |
| linear | 代码占比 <50% 的文档 |
