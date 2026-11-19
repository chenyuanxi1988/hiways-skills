---
name: alex-hiways-template
version: 2.0.0
description: 海华永泰律师事务所法律服务方案/法律文书标准模板。用于生成符合陈元熹律师团队正式法律文书格式的 Word 文档。触发词：海华模板、hiways template、法律服务方案、法律意见书、尽调报告、律师函、海华永泰格式。
repo: https://github.com/chenyuanxi1988/Skills
---

# Alex Hiways Template Skill

> GitHub 仓库：[chenyuanxi1988/Skills](https://github.com/chenyuanxi1988/Skills/tree/main/alex-hiways-template)

基于陈元熹律师团队实际交付的「某工程项目法律服务方案」提取的正式法律文书格式标准。

## 适用场景

- 法律服务方案 / 法律服务建议书
- 法律意见书
- 尽职调查报告 / 尽调资料清单
- 律师函
- 其他海华永泰正式对外法律文书

## 使用方法

用户说「用海华模板生成XX文件」或直接指定此 skill，然后提供内容（Markdown 文件或直接口述内容），AI 会自动生成格式化的 .docx 文件。

### 方式一：从 Markdown 文件生成（推荐）

```bash
python scripts/generate.py <输入.md> [输出.docx]
```

### 方式二：AI 直接生成

AI 根据用户描述的内容，使用 python-docx 库直接构造符合模板规范的文档。

## 格式规格

### 页面设置

- 纸张：A4（210mm × 297mm）
- 页边距：上下 2.54cm，左右 1.91cm
- 可用内容宽度：21.0 - 1.91 × 2 = **17.18cm** ≈ **9743 twips (dxa)**

### 字体规范

| 用途 | 西文字体 | 中文字体 | 字号 | 样式 |
|------|----------|----------|------|------|
| 正文 | FangSong (仿宋) | FangSong (仿宋) | 12pt | Regular |
| 一级标题 | FangSong | FangSong | 14pt | Bold |
| 二级标题 | FangSong | FangSong | 14pt | Bold |
| 三级标题 | FangSong | FangSong | 12pt | Bold |
| 封面标题 | FangSong | FangSong | 14pt | Bold, 居中 |
| 表头 | FangSong | FangSong | 9pt | Bold |
| 表格数据 | FangSong | FangSong | 8pt | Regular |
| 强调文字 | FangSong | FangSong | 12pt | Bold + #106F86 |

### 颜色体系

- 正文：黑色 #000000
- 强调色（slogan等）：青灰蓝 #106F86
- 表头底色：浅灰 #E8E8E8

### 段落格式

- 行距：正文 1.15 倍，封面标题 1.5 倍
- 对齐：封面居中，正文两端对齐（JUSTIFY），表单标签/元数据字段左对齐（LEFT）
- 首行缩进：中文正文 0.74cm（约 2 字符），表单标签/元数据字段无缩进
- 表单标签识别：含 `____`（填空）或以「标签：」开头的段落，自动设为 LEFT + 无缩进

### 封面结构（2页）

**第1页（封面）：**
1. 5 个空段落（为浮动 LOGO 留位）
2. LOGO：浮动锚定（anchored），6.47cm 宽，距左页边距 ~5.35cm（1926000 EMU）
3. 7 个空段落（浮动 LOGO 下方空间）
4. 标题行：14pt 加粗仿宋，居中，1.5倍行距
5. 2 个空段落
6. 「上海市海华永泰律师事务所」：14pt 加粗仿宋，居中
7. 5 个空段落
8. Slogan「专业赢得尊重，携手成就理想」：12pt 加粗 #106F86，居中，字间距 6twips
9. 2 个空段落
10. 城市列表：12pt 仿宋，居中，单个段落自然换行，字间距 -5twips

**第2页（信息页）：**
- 起草人/电话/邮箱/接收人/日期：左对齐，12pt 仿宋
- 主题：左对齐，12pt 加粗

### 页眉

- 「保密版本」水印文字，8pt 灰色，右对齐

### 表格格式

- **布局模式：fixed**（`tblLayout type="fixed"`），禁止 Word 自动撑宽
- **表格总宽：`tblW type="dxa"` 固定值**（不是 auto）
- 列宽配置（twips/dxa 单位）：

| 列数 | 各列宽度（twips） | 总宽 |
|------|-------------------|------|
| 6列 | 510, 907, 3289, 1021, 1588, 2313 | 9628 (~16.98cm) |
| 7列 | 510, 1134, 1985, 1021, 1021, 2155, 1917 | 9743 (~17.18cm) |

- 表头：9pt 加粗，浅灰底 #E8E8E8，居中
- 数据行：8pt，短列（序号、提供形态、备注）居中
- 单元格内边距：上下 20dxa，左右 40dxa

## 文件结构

```
alex-hiways-template/
├── SKILL.md                    ← 技能主文件（本文件）
├── EXTEND.md                   ← 用户自定义配置
├── assets/
│   └── hiways-logo.jpeg        ← 律所 LOGO
├── scripts/
│   └── generate.py             ← Markdown → DOCX 生成脚本
├── references/                 ← 参考文档（预留）
└── templates/                  ← 模板文件（预留）
```

## 依赖

```bash
pip install python-docx markdown beautifulsoup4 lxml
```

## ⚠️ 踩坑记录与关键注意事项

### 1. OOXML 单位体系：gridCol/tcW 必须用 twips (dxa)，不能用 EMU

**踩坑**：`w:gridCol w:w` 和 `w:tcW w:w` 的单位是 **twips (也叫 dxa)**，1cm ≈ 567 twips。python-docx 的 `Cm()` 返回的是 **EMU**（1cm = 360000 EMU），直接把 EMU 值写入 `gridCol` 会导致 Word 解析为天文数字。

| 单位 | 1cm 换算 | 适用属性 |
|------|---------|---------|
| twips/dxa | 567 | `gridCol w:w`, `tcW w:w`, `tblW w:w`, `spacing w:line` |
| EMU | 360000 | `cell.width`（python-docx API）, `posOffset`（图片定位） |
| half-points | 2 | 字号 `sz w:val`（12pt = 24） |

**正确做法**：
```python
# gridCol 和 tcW 用 twips
COL_WIDTHS_TWIPS = {6: [510, 907, 3289, 1021, 1588, 2313]}

# cell.width 用 EMU（python-docx API 内部转换）
width_emu = [tw * 635 for tw in width_twips]  # 1 twip = 635 EMU
cell.width = width_emu[ci]
```

### 2. 表格必须同时设置 tblW + tblLayout=fixed + gridCol

**踩坑**：只设置 `tblLayout=fixed` 而不设 `tblW`，Word 仍会按 auto 模式计算宽度，导致列宽被忽略。

**必须三件套**：
```python
# 1. tblW 设为固定 dxa 值
tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:type="dxa" w:w="{total_twips}"/>'))
# 2. tblLayout 设为 fixed
tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
# 3. gridCol 设定各列宽度
gc.set(qn('w:w'), str(width_twips[i]))
```

### 3. 封面 LOGO 需用浮动锚定（anchor），不能 inline

**踩坑**：inline 图片受行高（`w:spacing w:lineRule="exact"`）限制，LOGO 会被裁剪。即使设了 exact 行高 1846 twips（参考文件值），仍可能有部分被遮挡。

**正确做法**：将 inline 图片转为 `wp:anchor`（浮动锚定），这样图片独立于段落行高：
- 水平定位：`positionH relativeFrom="margin"` + `posOffset` 1926000 EMU（~5.35cm）
- 垂直定位：`positionV relativeFrom="paragraph"` + `posOffset` 0
- 宿主段落设极小行高（240 twips exact），不占多余空间

### 4. 封面标题需过滤律所名称，避免重复

**踩坑**：Markdown 源文件的 `#` 标题常含「海华永泰律师事务所」/「HIWAYS LAW FIRM」，而模板会自动生成律所名，导致封面出现两次。

**正确做法**：`filtered_titles = [t for t in title_lines if not any(kw in t for kw in firm_keywords)]`

### 5. 城市列表必须是一个段落，不要拆分

**踩坑**：将城市列表拆成3个段落会导致 Word 中排版不一致。应使用单个段落 + 负字间距（`w:spacing w:val="-5"`）让内容自然换行。

### 6. 封面空段落数量需要精确调校

封面内容必须在一页内完成。空段落数量的调整需综合考虑：
- LOGO 浮动后，其宿主段落几乎不占空间，需要额外的空段落把后续文字推到正确位置
- 城市列表不能溢出到第2页
- 当前调校值：5（顶部）+ 7（LOGO后）+ 2（标题后）+ 5（律所名后）+ 2（slogan后）

### 7. WD_BREAK_PAGE 导入路径

**踩坑**：`from docx.enum.text import WD_BREAK_PAGE` 在某些版本不存在。

**正确做法**：`from docx.enum.text import WD_BREAK`，然后使用 `WD_BREAK.PAGE`。

### 8. 对外发送文件需删除内部备注

Markdown 源文件中可能含 `【注：成本类，由合作机构实施...】` 等内部协作备注，正式交付版必须清除。generate.py 目前不做自动过滤，需在 Markdown 源文件中手动删除或预处理。

### 9. 表单标签/元数据字段必须左对齐、无首行缩进

**踩坑**：文档末尾的表单类段落（如「交接日期：2026年__月__日」「移交人声明：」「移交人签名：____」）被当作普通正文处理，加上 JUSTIFY 对齐和首行缩进，导致排版不合规。

**正确做法**：通过 `is_form_label()` 识别表单类段落，设为 LEFT 对齐、无缩进。识别规则：
1. 含4个以上连续下划线（`____`）→ 填空项
2. 以「标签：」或「**标签：**」开头 → 元数据字段

```python
def is_form_label(text):
    if re.search(r'_{4,}', text):
        return True
    stripped = text.strip()
    if re.match(r'^\*{0,2}[^*]+\*{0,2}[：:]', stripped):
        return True
    return False
```

在段落渲染逻辑中：`is_heading_like → 加粗` > `is_form_label → 左对齐无缩进` > `普通正文 → JUSTIFY + 缩进`。

### 10. Markdown 横线分隔符（---）不渲染到 DOCX

**踩坑**：Markdown 中的 `---` 横线被渲染为长灰色横线（`─` × 50），插入每个章节之间，视觉上过于突兀，不符合正式法律文书风格。

**正确做法**：`horizontal_rule` 类型直接跳过（`pass`），不生成任何内容。文档末尾署名区的分隔线由 `add_footer_info()` 单独处理，不受此影响。

```python
elif t == 'horizontal_rule':
    # Skip: 不在章节间插入分隔横线
    pass
```

## 注意事项

1. 生成文档后，建议用 Word 打开微调分页和图片位置
2. 表格列宽已通过 fixed layout + dxa 单位控制，无需手动调整
3. 此模板适用于正式对外法律文书，内部备忘录可使用简化版
4. 如需新增表格列数配置，在 `COL_WIDTHS_TWIPS` 字典中添加对应条目
