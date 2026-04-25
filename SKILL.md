---
name: docx-to-md
description: "Convert Word documents (.docx) to Markdown format, including image extraction and large file support. Use when the user wants to convert a .docx file to Markdown."
---

# Word 转 Markdown（大文件支持版）

将 `.docx` 文件转换为 Markdown 格式，支持图片提取、分块写入、大文档处理。

## Python 路径（Windows）

```
C:/Users/Lecoo/AppData/Local/Programs/Python/Python310/python.exe
```

## 输出结构

```
<文件名>_md/
├── output.md       # 主 Markdown 文件
└── images/         # 提取的图片（image_001.png ...）
```

输出目录 = 输入文件所在目录 + `/<文件名（去掉.docx）>_md`

## 依赖

已安装：`markitdown[docx]`、`python-docx`

## 工作流

### Step 1：确认输入

从用户消息提取 `.docx` 文件路径，不明确时询问。

### Step 2：写出转换脚本

将以下 Python 脚本写入 `C:/Users/Lecoo/AppData/Local/Temp/docx_convert.py`。
**必须用脚本文件方式，不要用 `-c` 内联模式**（避免大文件时命令行长度限制）。

使用 Write 工具写入前 50 行，再用 Edit 追加剩余内容：

**前半（Write 写入）**：
```python
import os, sys, zipfile, shutil
from pathlib import Path
from markitdown import MarkItDown

def convert(input_path, output_dir):
    input_path = Path(input_path)
    output_dir = Path(output_dir)
    images_dir = output_dir / 'images'
    output_dir.mkdir(parents=True, exist_ok=True)
    images_dir.mkdir(exist_ok=True)

    print(f'[1/4] 读取文档：{input_path.name}，大小 {input_path.stat().st_size/1024:.1f} KB')

    print('[2/4] 提取图片...')
    image_count = 0
    img_map = {}
    with zipfile.ZipFile(input_path, 'r') as z:
        media_files = [f for f in z.namelist() if f.startswith('word/media/')]
        for i, media_file in enumerate(sorted(media_files), 1):
            ext = Path(media_file).suffix.lower()
            new_name = f'image_{i:03d}{ext}'
            dest = images_dir / new_name
            with z.open(media_file) as src, open(dest, 'wb') as dst:
                shutil.copyfileobj(src, dst)
            img_map[Path(media_file).name] = new_name
            image_count += 1
    print(f'      提取图片：{image_count} 张')
```

**后半（Edit 追加）**：
```python
    print('[3/4] 转换文本结构...')
    md_engine = MarkItDown()
    result = md_engine.convert(str(input_path))
    md_text = result.text_content

    if image_count > 0:
        img_section = '\n\n---\n\n## 提取的图片\n\n'
        for orig, new in img_map.items():
            img_section += f'![{orig}](images/{new})\n\n'
        md_text += img_section

    print('[4/4] 写入输出文件...')
    output_file = output_dir / 'output.md'
    with open(output_file, 'w', encoding='utf-8') as f:
        for i in range(0, len(md_text), 65536):
            f.write(md_text[i:i+65536])

    size_kb = output_file.stat().st_size / 1024
    print(f'\n转换完成！')
    print(f'  输出文件：{output_file}')
    print(f'  Markdown 大小：{size_kb:.1f} KB')
    print(f'  提取图片：{image_count} 张')
    print(f'  字符总数：{len(md_text):,}')

if __name__ == '__main__':
    convert(sys.argv[1], sys.argv[2])
```

### Step 3：执行脚本

```bash
C:/Users/Lecoo/AppData/Local/Programs/Python/Python310/python.exe \
  C:/Users/Lecoo/AppData/Local/Temp/docx_convert.py \
  "<INPUT_PATH>" \
  "<OUTPUT_DIR>"
```

### Step 4：验证输出

用 Read 工具读取 `output.md` 前 30 行，确认标题、段落、表格结构正确。

### Step 5：报告结果

告知用户输出路径、Markdown 大小、提取图片数量及输出目录结构。

## 备用方案（markitdown 失败时）

用 python-docx 手动提取：

```python
from docx import Document
from pathlib import Path

HEADING_MAP = {
    'Heading 1': '#', 'Heading 2': '##', 'Heading 3': '###',
    'Heading 4': '####', 'Heading 5': '#####',
    '标题 1': '#', '标题 2': '##', '标题 3': '###',
}

def table_to_md(table):
    rows = []
    for i, row in enumerate(table.rows):
        cells = [c.text.strip().replace('\n', ' ') for c in row.cells]
        rows.append('| ' + ' | '.join(cells) + ' |')
        if i == 0:
            rows.append('| ' + ' | '.join(['---'] * len(cells)) + ' |')
    return '\n'.join(rows)

def convert_fallback(input_path, output_dir):
    doc = Document(input_path)
    lines = []
    for block in doc.element.body:
        tag = block.tag.split('}')[-1]
        if tag == 'p':
            from docx.text.paragraph import Paragraph
            para = Paragraph(block, doc)
            text = para.text.strip()
            if not text:
                lines.append('')
                continue
            prefix = HEADING_MAP.get(para.style.name, '')
            if prefix:
                lines.append(f'{prefix} {text}')
            elif 'List' in para.style.name:
                lines.append(f'- {text}')
            else:
                lines.append(text)
            lines.append('')
        elif tag == 'tbl':
            from docx.table import Table
            tbl = Table(block, doc)
            lines.append(table_to_md(tbl))
            lines.append('')
    output_file = Path(output_dir) / 'output.md'
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f'备用方案完成：{output_file}')
```

## 注意事项

- 图片存放在 `images/` 子文件夹，Markdown 中使用相对路径引用
- 密码保护的文档无法转换，需用户先去除保护
- 批注和修订记录不会被保留
- 中文文档完全支持
