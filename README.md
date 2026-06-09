# Markdown Style Converter
# Markdown 样式转换器

Convert Markdown to styled Word or HTML with one click.
一键将 Markdown 转换为带精美样式的 Word 或 HTML。

8 professional styles for AI-generated content.
专为 AI 产出内容设计，8 种专业版式。

Word tables render beautifully, especially for bilingual documents.
Word 表格显示效果极佳，尤其适合双语文档。

## Install / 安装

```bash
git clone https://github.com/shirleyharleywiley/markdown-to-multi-style.git ~/.claude/skills/mdstyle
```

## 8 Styles / 8 种版式

| 版式Style | 适合输出格式Color | 典型适用场景Best For |
| ------| ------| ------|  
| block| html | 简历 resume |
| deco | html | 技术博客 tech blog |
| pill | html | 公众号文章 article |
| stripe | html, word | API文档 document |
| cmd | html, word | 报告论文 paper |
| notion | html，word | 双语内容 multi-language content |
| medium | html, word | 超长文 long content |
| Linear | html | 代码密集 code |

## Quick Start / 快速开始

```bash
/mdstyle /path/to/article.md
```

AI analyzes the file and recommends the best style.
AI 自动分析文件名和内容，推荐最适合的版式。

## Tips / 避坑指南

| Style | Avoid Using For |
|-------|-----------------|
| gradient-pill | 含表格的文档 / Docs with tables |
| bottom-deco | 正式工作报告 / Formal reports |
| left-block | 标题密集的文档 / Title-heavy docs |
| linear | 代码占比 <50% 的文档 / Code <50% |
