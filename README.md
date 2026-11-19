# Hiways Skills · 海华永泰文书模板

供海华永泰律师及协作人员使用的通用 AI Skills。

## 当前技能

- [alex-hiways-template](./alex-hiways-template/SKILL.md)：将 Markdown 法律文书转换成统一格式的 Word（DOCX）。
- 模板格式、Logo 与转换脚本位于同一技能目录中。

## 最简单的使用方法

1. 在支持文件型 Skills 的 AI 客户端中，安装整个 `alex-hiways-template` 文件夹（不要只复制 `SKILL.md`）。
2. 告诉 AI：“使用海华模板将这份 Markdown 生成 Word 文档”。
3. 需要手动转换时，先安装依赖：
   `pip install python-docx markdown beautifulsoup4 lxml`
4. 运行：
   `python alex-hiways-template/scripts/generate.py input.md output.docx`

不同客户端的 Skill 安装方式可能不同，须按其实际支持的目录或导入机制安装。

## 维护规则

- **唯一维护源**：私有 `chenyuanxi1988/Skills` 仓库中的 `alex-hiways-template`。
- **本仓库是脱敏发布副本**：请不要直接在这里修改模板规则或脚本，以免两个版本产生分叉。
- 发布前检查客户名称、项目识别信息、内部备注、私密联系方式和密钥。
- 遵循 `SKILL.md` 中的 `version` 字段；正式发布和自动同步机制尚在建设中。
- 文件可以与私有源保持相同哈希，但这**不代表**所有配置都适合公开发布：必须先完成敏感信息审核。

## 使用前注意

脚本目前**不会自动过滤**源 Markdown 中的内部备注；对外发送前，请人工检查生成的 DOCX，确认无客户保密信息、内部批注或排版异常。

---
作者：陈元熹律师 · 上海市海华永泰律师事务所
