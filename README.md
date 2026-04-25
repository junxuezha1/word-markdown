# docx-to-md

> 一个为 Claude Code 设计的 Skill，将 Word 文档（`.docx`）转换为结构清晰的 Markdown 文件，支持大文件、图片提取与备用解析引擎。

---

## 功能特性

| 特性 | 说明 |
|------|------|
| **标题层级保留** | `Heading 1~5` 自动映射为 `#` ~ `#####` |
| **表格转换** | Word 表格转为标准 Markdown 管道符格式 |
| **列表支持** | 有序 / 无序列表完整保留 |
| **图片提取** | 从 docx ZIP 结构中提取嵌入图片，保存至 `images/` 子目录，并在 Markdown 中自动引用 |
| **大文件支持** | 脚本文件执行模式 + 64 KB 分块写入，避免命令行长度限制与内存峰值 |
| **双引擎架构** | 主选 `markitdown[docx]`（微软出品），失败时自动切换 `python-docx` 手动解析 |
| **中文全兼容** | 中文文件名、中文内容、中文路径均完整支持 |

---

## 输出结构

```
<原文件名>_md/
├── output.md        ← 转换后的 Markdown 主文件
└── images/
    ├── image_001.png
    ├── image_002.jpg
    └── ...
```

---

## 使用方式

在 Claude Code 对话中，直接告诉 Claude：

```
把 /path/to/document.docx 转换成 Markdown
```

Claude 会自动触发本 Skill，在原文件同目录生成 `<文件名>_md/` 输出文件夹。

---

## 依赖

```bash
pip install "markitdown[docx]"
pip install python-docx
```

---

## 转换示例

**输入**：一份 241 KB 的法学研究课题申请书（含表格、图片、多级标题）

**输出**：251 KB Markdown 文本，707 行，提取图片 2 张，耗时 < 3 秒

```markdown
**湖南省法学会** **2025** **年度研究课题**

**结项申请书**

# 第一部分 课题基本情况

| 课题名称 | "三高四新"发展战略背景下长株潭都市圈数据协同的法治保障研究 |
| --- | --- |
| 立项编号 | 24HNFX-C-004 |
| 资助经费总额 | 4 万元 |
...
```

---

## 安装为 Claude Code Skill

将 `SKILL.md` 放入你的 Claude Code skills 目录：

```
~/.claude/skills/docx-to-md/SKILL.md
```

重启 Claude Code 后即可使用。

---

## 注意事项

- 密码保护的文档需先解除保护
- 批注与修订记录不会被保留（仅保留接受后正文）
- 图片在正文中的精确位置取决于原文档的行内图片锚定方式；若 markitdown 无法定位，图片统一追加至文末

---

## License

MIT
