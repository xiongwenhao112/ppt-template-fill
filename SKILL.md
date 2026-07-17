---
name: ppt-template-fill
description: 基于用户提供的 PPT 模板制作演示文稿时使用。解析任意 .pptx/.potx 模板(含无占位符的自由文本框模板)每页的可填充槽位，克隆模板页并在完整保留原版式、字体、配色的前提下填入新内容，输出与模板视觉一致的原生可编辑 PPTX。触发语如：用这个模板做 PPT、按公司模板生成、套模板写 PPT。
---

# PPT Template Fill

把用户的内容需求"套"进用户指定的 PPTX 模板：模板的每一页都是候选版式，克隆整页后只替换文字/图片/表格内容，版式、字体、配色、装饰元素一律不动。产物是原生可编辑的 PPTX，不经过 HTML 中间态。

`SKILL.md` 所在目录记为 `<skill-root>`。

## 环境要求

- Python 3.10+，`python-pptx`（必需；缺失时 `pip install python-pptx`）
- 截图(可选增强)：Windows 本机 PowerPoint，或 LibreOffice + PyMuPDF。都没有时跳过截图步骤，仅凭 manifest 工作并在交付时说明未做视觉复核。

## 命令速查

```bash
# 1. 模板槽位清单(先 --compact 看全貌, 再 --out 存完整契约)
python <skill-root>/scripts/inspect_template.py <模板.pptx> --compact
python <skill-root>/scripts/inspect_template.py <模板.pptx> --out <workdir>/manifest.json

# 2. 每页截图(选版式和成品 QA 都用它)
python <skill-root>/scripts/snapshot.py <某.pptx> <输出目录> [--width 1440]

# 3. 校验计划 / 生成成品
python <skill-root>/scripts/build_deck.py <workdir>/plan.json --check
python <skill-root>/scripts/build_deck.py <workdir>/plan.json
```

plan.json 的完整格式和槽位 ID 规则见 `<skill-root>/references/plan-format.md`。

## 工作流

1. **确认输入**：模板文件路径 + 内容需求(主题、受众、页数、重点)。用户没给模板路径时先问；用户只给了素材文档时，从文档中提炼内容。输出目录用当前会话工作目录下的 `output/<deck-name>/`，不要写进 `<skill-root>`。
2. **解析模板**：先 `--compact` 看每页角色和槽位概况，再 `--out` 保存完整 manifest。随后对模板跑 `snapshot.py`，用 Read 逐张查看截图——截图是判断每页真实用途和视觉密度的第一依据，`role_guess` 只是提示。
3. **规划结构**：先定章节大纲，再为成品每一页选模板源页。同一源页可以多次使用(这是常态：正文页会被反复克隆)。封面用封面页、结尾用结尾页，各用一次。目录页条目数以模板实际槽位为准，反推章节数量，保证目录与章节页一一对应。
4. **写 plan.json**：每页 `source`(模板页码) + `texts`/`images`/`notes`。遵守下方填写规则。
5. **校验**：`build_deck.py plan.json --check`。E(错误)必须修复；W(警告)逐条评估——超预算的文本先精简文案，确实无法精简再接受并在 QA 时重点看该页。
6. **生成 + QA**：去掉 `--check` 生成成品，对成品跑 `snapshot.py`，逐页 Read 检查三件事：文字溢出/遮挡、模板演示文案残留、图片变形错位。发现问题改 plan 重新生成，最多 2 轮；仍有问题时如实说明。
7. **交付**：给出成品 pptx 绝对路径，说明用了模板哪几页、做了哪些假设。若有跳过的步骤(如无法截图)一并说明。

## 填写规则

- **全量覆写**：选定源页后，manifest 里该页所有 `decorative: false` 且有样本文字的文本槽**必须全部**写进 `texts`。漏填的槽会原样残留模板演示文案，属于交付失败。`--check` 会对漏填逐条警告。
- **装饰槽不动**：`decorative: true`(页码、序号、LOGO 字样等)一律不填。
- **字数纪律**：不超过槽位 `max_chars`；没有 `max_chars` 时以 `sample_len` 为基准上下浮动 20%。标题槽只写短语不写句子。空占位符看 `layout_prompt` 判断用途。宁可精炼，不可溢出——python-pptx 不会自动缩字号。
- **分段**：`\n` 分段，段落数尽量与样本 `paragraphs` 一致；`max_lines: 1` 的槽不要多段。
- **内容语言跟随用户**：用户用英文沟通就写英文文案，模板残留语言不算数。
- **图片槽**：用户给了素材才替换(路径相对 plan.json，一张素材只用一次)；没给素材时保留模板原图——原图与模板风格协调，不要伪造路径或强行配图。替换时引擎会按槽位宽高比自动居中裁切，选图时留意 manifest 中的 `aspect`。
- **表格**：按 `<shape_id>!r行c列` 逐格填写，行列数以 manifest 为准，不能增删行列。
- **不碰样式**：本 skill 只填内容。不要试图改颜色、字号、位置、增删形状；模板长什么样，成品就长什么样。
- **图表/SmartArt 页**（manifest `other` 列出）：当前版本不能改其数据。避免选这类页；用户坚持要用时保留原样并明确告知。
- **备注加分项**：每页 `notes` 可写 1-2 句演讲备注，整体委托时默认写上。

## 常见陷阱

- 模板页数少、正文版式单一是模板本身的限制，不要为凑页数硬造视觉——多复用现有正文页，靠内容分段。
- 目录条目多于实际章节时，多余条目填入真实的收尾性内容(如"总结与展望")，不许留"内容简述"这类样本。
- 校验通过≠视觉合格：CJK 长词、全角标点都会造成意外换行，成品截图必须过目。
