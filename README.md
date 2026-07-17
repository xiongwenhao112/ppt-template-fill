# ppt-template-fill

![License: MIT](https://img.shields.io/badge/License-MIT-blue?style=flat-square)
![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=flat-square)
![Claude Code](https://img.shields.io/badge/Claude%20Code-Supported-6B5B95?style=flat-square)
![Codex](https://img.shields.io/badge/Codex-Supported-222222?style=flat-square)

给 AI Agent（Claude Code / Codex 等）用的 PPT Skill：**基于你自己的 PPTX 模板生成演示文稿**。

*An agent skill that fills YOUR PowerPoint template: parses any .pptx as a layout library, clones template slides, and fills in new content while preserving 100% of the original design.*

把模板和内容需求丢给 Agent，一句话生成一份**和模板视觉完全一致、原生可编辑**的 PPTX：

> 用 `D:\模板\公司模板.pptx` 这个模板，帮我做一份《2026 上半年销售复盘》的 PPT，10 页左右，面向管理层。

## 它解决什么问题

市面上的 PPT skill（如 dashi-ppt-skill）大多用**自带**的预置主题库，风格再漂亮也不是你的——公司模板、学校课件模板、比赛指定模板这类"必须长这样"的场景覆盖不了。

本 skill 反过来：**你的模板就是版式库**。而且专门处理了真实世界模板的痛点——市面上大量中文模板**根本不用标准占位符**，全靠自由文本框和组合形状排版，常规"往占位符里填字"的方案直接失效。本 skill 以整页克隆 + 文字 run 级替换工作，对这类模板同样有效。

## 原理

模板的每一页都是候选版式。Agent 先解析模板得到每页的"槽位契约"（哪些文本框可以填、字数预算多少、哪些是装饰不能动、图片槽宽高比如何），再把内容规划写成受约束的 `plan.json`，最后由生成器克隆模板页、只替换内容：

```
inspect_template.py  →  模板槽位 manifest + 逐页截图
        ↓
   plan.json         →  每页: 选模板哪一页 + 各槽位填什么
        ↓
build_deck.py        →  校验(字数/漏填/路径) → 克隆页 → 填充 → 成品 pptx
        ↓
snapshot.py          →  成品逐页截图, Agent 目检后交付
```

设计哲学（借鉴 dashi-ppt-skill 的方法论）：LLM 不擅长稳定的视觉设计，但很擅长填结构化 JSON。所以把 AI 的自由度压缩到只剩"选版式"和"写文案"，错误要么被契约挡在生成前，要么被截图抓在交付前。

## 特性

- **版式 100% 保真**：整页克隆（含图片/超链接关系重映射），只动文字 run，字体、字号、加粗、颜色、装饰元素全部继承
- **产物即原生 PPTX**：无 HTML 中间态，PowerPoint 直接打开编辑
- **吃"野生"模板**：无占位符、自由文本框、组合嵌套、多母版都能处理
- **槽位契约**：自动识别可填文本槽（含字数预算）、装饰槽（页码/LOGO，防误填）、图片槽（宽高比）、表格、空占位符（附版式用途提示）
- **图片替换**：按槽位宽高比自动居中裁切（srcRect，不重编码不变形）
- **表格逐格填写**、**演讲者备注**写入
- **双重质检**：生成前校验（模板文案残留、超字数、坏路径），生成后 PowerPoint COM 逐页截图供 Agent 目检

## 环境要求

- Python 3.10+，`python-pptx`
- 截图（可选增强）：Windows + PowerPoint，或 LibreOffice + PyMuPDF；没有时跳过截图纯靠契约工作

## 安装

```powershell
# Windows PowerShell
pip install python-pptx
git clone https://github.com/xiongwenhao112/ppt-template-fill.git "$env:USERPROFILE\.claude\skills\ppt-template-fill"
```

```bash
# macOS / Linux
pip install python-pptx
git clone https://github.com/xiongwenhao112/ppt-template-fill.git ~/.claude/skills/ppt-template-fill
```

其他 Agent：把本目录放到对应技能目录，或直接让它读 `SKILL.md`。

## 使用

装好后对 Agent 说：

> 用 `xxx.pptx` 这个模板，帮我做一份 XX 主题的 PPT

也可以丢一份 Word/Markdown 素材让它提炼内容。改稿时直接说"第 3 页标题换成……"，或者用 PowerPoint 打开成品自己改。

完整教程（含手动命令行用法、manifest 字段速查、FAQ）见 **[使用教程.md](使用教程.md)**。

## 目录结构

```
ppt-template-fill/
├── SKILL.md                  # Agent 工作流与填写纪律
├── 使用教程.md                # 人类用户教程
├── scripts/
│   ├── pptxkit.py            # 核心库: 槽位识别 / 整页克隆 / 保格式文本替换
│   ├── inspect_template.py   # 模板 → 槽位契约 manifest
│   ├── build_deck.py         # plan.json → 成品 pptx(含校验)
│   ├── snapshot.py           # pptx → 逐页 PNG(PowerPoint COM / LibreOffice)
│   └── snapshot.ps1
└── references/
    └── plan-format.md        # plan.json 与槽位 ID 规范
```

## 已知限制

- 图表（Chart）与 SmartArt 数据暂不可改（保留原样）
- 表格行列数固定，只能改格子内文字
- 文本超长不会自动缩字号，靠字数预算 + 成品截图目检兜底
- 不做任何样式自定义——模板长什么样，成品就长什么样（这是特性）

## Roadmap

- [ ] 图表数据填充（cat/val 缓存重写）
- [ ] 文本超预算时自动缩字号 / 自动分页
- [ ] `.potx` 模板与多模板库管理
- [ ] LibreOffice 截图链路的开箱即用打包

## License

[MIT](LICENSE)
