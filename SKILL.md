---
name: mdstyle
description: 将md文件转换成html文件或word文件。当用户指定了一个或多个.md文件，并且提到“转html”、“转word”、“生成公众号文章”、“生成报告”时调用。
---

# Markdown 样式转换器

## 使用方法

根据用户指定的输出文件格式及AI分析文件名和内容 - 推荐版式-确认后生成。

### step1：根据用户指定的输出文件格式及AI分析文件名和内容后得出的场景，选出最优版式

如果用户指定了明确的版式名称
| 版式 | 适合输出格式 | 典型适用场景 |
| ------| ------| ------|  
| block| html | 简历 |
| deco | html | 技术博客 |
| pill | html | 公众号文章 |
| stripe | html, word | API文档 |
| cmd | html, word | 报告论文 |
| notion | html，word | 双语内容 |
| medium | html, word | 超长文 |
| Linear | html | 代码密集 |

### step2：将step1选出的版式推荐给用户，让用户确认

如果用户提及“全部”或“都要”，则8种版式的对应文件都生成一遍。

### step3：生成对应文件

根据用户确认的版式
Run 'python scripts/converter.py {filename} [html|docx] [style]' 生成对应文件

例：生成 HTML 'python scripts/converter.py article.md html block'
例：生成 Word 'python scripts/converter.py article.md docx stripe'

### 避坑指南

| 版式| 避免用于|
| ------ | ------ | 
| pill | 含表格的文档 |
| deco | 正式工作报告 |
| block | 标题密集的文档 |
| Linear | 代码占比 <50% |

> 查看 ‘references/examples/’ 目录查看每种版式的实际效果