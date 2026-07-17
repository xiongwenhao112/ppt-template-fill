---
name: ppt-template-fill
description: 基于用户提供的 PPT 模板制作演示文稿时使用。解析任意 .pptx/.potx 模板(含无占位符的自由文本框模板)每页的可填充槽位，克隆模板页并在完整保留原版式、字体、配色的前提下填入新内容，输出与模板视觉一致的原生可编辑 PPTX。触发语如：用这个模板做 PPT、按公司模板生成、套模板写 PPT。
---

# PPT Template Fill

把用户的内容需求"套"进用户指定的 PPTX 模板：模板的每一页都是候选版式，克隆整页后只替换文字/图片/表格内容，版式、字体、配色、装饰元素一律不动。产物是原生可编辑的 PPTX，不经过 HTML 中间态。

`SKILL.md` 所在目录记为 `<skill-root>`。

## 红线(先读。违反任何一条 = 交付失败)

1. **只能用本 skill 的脚本操作 PPTX**。解析、克隆、填充、截图全部通过 `inspect_template.py` / `build_deck.py` / `snapshot.py` 完成。**禁止自己编写或执行任何 python-pptx / OOXML / 其他 PPT 读写代码**——网上流传的克隆代码会产出畸形文件(如一页挂两个版式关系)，你手写的填充不会保留格式。
2. **脚本缺失或报错时：停下，把报错原样告诉用户**。手写替代实现不是备选项；缺 `python-pptx` 就先 `pip install python-pptx`。
3. **禁止新增/删除/移动任何形状**。不许新建文本框，不许改字体、字号、颜色、位置。你的全部自由度只有两个：成品每一页选模板哪一页 + 每个槽位填什么文字。标题、副标题、页眉都在模板槽位里，不存在"模板没给标题所以我加一个"的情况——没有该槽位就说明这页版式不需要它。
4. **选中的模板页，非装饰文本槽必须全部填**(plan 里不许残留任何「【待填】」)。想保留模板原文，就把原文作为值填回去。漏填的槽会原样残留"内容简述"之类演示文案，属于最难看的交付事故，`--check` 默认按错误阻断。
5. **用户给了素材文档时，标题/数字/术语/结论必须出自素材**。禁止填入素材里没有的泛化文案(如素材讲"根因诊断"却写成"趋势预测")。

## 环境要求

- Python 3.10+，`python-pptx`(必需；缺失时 `pip install python-pptx`)
- 截图(可选增强)：Windows 本机 PowerPoint，或 LibreOffice(+PyMuPDF)。用 `snapshot.py --probe` 探测；不可用时按"无截图 QA"降级，并在交付时说明未做视觉复核。

## 命令速查

```bash
# 0. 自检: 脚本可达 + 截图后端探测
python <skill-root>/scripts/inspect_template.py --help
python <skill-root>/scripts/snapshot.py --probe

# 1. 模板槽位清单(先 --compact 看全貌, 再 --out 存完整契约)
python <skill-root>/scripts/inspect_template.py <模板.pptx> --compact
python <skill-root>/scripts/inspect_template.py <模板.pptx> --out <workdir>/manifest.json

# 2. 每页截图(选版式和成品 QA 都用它)
python <skill-root>/scripts/snapshot.py <某.pptx> <输出目录> [--width 1440]

# 3. 生成计划骨架(选定页序后, 每个非装饰槽预填「【待填】+ 原文」)
python <skill-root>/scripts/build_deck.py <workdir>/plan.json --scaffold \
    --template <模板.pptx> --pages 1,2,4,4,4,6 --deck-out <workdir>/成品.pptx

# 4. 校验计划 / 生成成品
python <skill-root>/scripts/build_deck.py <workdir>/plan.json --check
python <skill-root>/scripts/build_deck.py <workdir>/plan.json
```

plan.json 的完整格式和槽位 ID 规则见 `<skill-root>/references/plan-format.md`。

## 工作流

0. **自检**：跑上面两条自检命令。`inspect_template.py --help` 失败说明脚本不可达——停止并告知用户(红线 2)。`--probe` 失败则记下"本次无视觉复核"，后续走无截图 QA。
1. **确认输入**：模板文件路径 + 内容需求(主题、受众、页数、重点)。用户没给模板路径时先问；用户只给了素材文档时，从文档中提炼内容(红线 5)。输出目录用当前会话工作目录下的 `output/<deck-name>/`，不要写进 `<skill-root>`。
2. **解析模板**：先 `--compact` 看每页角色和槽位概况，再 `--out` 保存完整 manifest。随后对模板跑 `snapshot.py`，用 Read 逐张查看截图——截图是判断每页真实用途和视觉密度的第一依据，`role_guess` 只是提示。
3. **规划结构**：先定章节大纲，再为成品每一页选模板源页。同一源页可以多次使用(这是常态：正文页会被反复克隆)。封面用封面页、结尾用结尾页，各用一次。目录页条目数以模板实际槽位为准，反推章节数量，保证目录与章节页一一对应。
4. **生成骨架**：按选好的页序跑 `--scaffold`。骨架里每个非装饰文本槽都带「【待填】+ 模板原文」，原文就是字数和分段的参照。
5. **填空**：把每个「【待填】」值整体替换成新文案(保留模板原文 = 只删掉「【待填】 」前缀)。遵守下方填写规则。表格格子已按原文预填，只改需要变的。图片槽见 `_image_slots` 提示，有素材才填 `images`。
6. **校验**：`build_deck.py plan.json --check`。E(错误)必须修复——「未填将残留」「仍含待填」都是 E；W(警告)逐条评估——超预算的文本先精简文案，确实无法精简再接受并在 QA 时重点看该页。
7. **生成 + QA**：去掉 `--check` 生成成品。**有截图后端**：对成品跑 `snapshot.py`，逐页 Read 检查三件事——文字溢出/遮挡、模板演示文案残留、图片变形错位；发现问题改 plan 重新生成，最多 2 轮。**无截图后端**：把 `--check` 的警告清单逐条核对(尤其超预算槽位的文案再精简一轮)，交付时明确说明未做视觉复核。
8. **交付**：给出成品 pptx 绝对路径，说明用了模板哪几页、做了哪些假设、QA 方式(截图目检 or 契约校验)。

## 填写规则

- **字数纪律**：不超过槽位 `max_chars`(见 manifest 或骨架 `_hints`)；没有 `max_chars` 时以样本原文长度为基准上下浮动 20%。标题槽只写短语不写句子。空占位符看 `layout_prompt` 判断用途。宁可精炼，不可溢出——python-pptx 不会自动缩字号。
- **分段**：`\n` 分段，段落数尽量与样本一致；`max_lines: 1` 的槽不要多段。
- **装饰槽不动**：`decorative: true`(页码、序号、LOGO 字样等)骨架里不会出现，也不要手动加进 plan。
- **内容语言跟随用户**：用户用英文沟通就写英文文案，模板残留语言不算数。
- **图片槽**：用户给了素材才替换(路径相对 plan.json，一张素材只用一次)；没给素材时保留模板原图——原图与模板风格协调，不要伪造路径或强行配图。替换时引擎会按槽位宽高比自动居中裁切，选图时留意 `aspect`。
- **表格**：按 `<shape_id>!r行c列` 逐格填写，行列数以 manifest 为准，不能增删行列。
- **图表/SmartArt 页**(manifest `other` 列出)：当前版本不能改其数据。避免选这类页；用户坚持要用时保留原样并明确告知。
- **备注加分项**：每页 `notes` 可写 1-2 句演讲备注，整体委托时默认写上。

## 常见陷阱

- 模板页数少、正文版式单一是模板本身的限制，不要为凑页数硬造视觉——多复用现有正文页，靠内容分段。**绝不允许**通过自己加文本框来"补"版式(红线 3)。
- 目录条目多于实际章节时，多余条目填入真实的收尾性内容(如"总结与展望")，不许留"内容简述"这类样本。
- 校验通过≠视觉合格：CJK 长词、全角标点都会造成意外换行，能截图时成品必须过目。
- 发现自己正在写 `from pptx import ...` 或解压 pptx 改 XML——立即停止，回到脚本(红线 1)。
