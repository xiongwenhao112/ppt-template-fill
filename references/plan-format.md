# plan.json 格式

```json
{
  "template": "模板.pptx 路径(绝对路径, 或相对 plan.json)",
  "output": "成品.pptx 路径(绝对路径, 或相对 plan.json)",
  "slides": [
    {
      "source": 3,
      "texts": {
        "14": "新标题\n第二段文字",
        "9!r2c3": "表格第2行第3列的内容"
      },
      "images": {
        "22": "assets/architecture.png"
      },
      "notes": "本页演讲备注(可选)"
    }
  ]
}
```

## 字段说明

| 字段 | 必填 | 说明 |
|------|------|------|
| `template` | 是 | 模板文件。生成时模板本身不会被修改 |
| `output` | 是 | 成品输出路径，目录会自动创建 |
| `slides[].source` | 是 | 模板页码(1-based)。**可以重复**——正文页通常被多个成品页复用 |
| `slides[].texts` | 否 | 槽位 ID → 新文本。`\n` 表示分段 |
| `slides[].images` | 否 | 图片槽 ID → 图片文件路径(相对 plan.json 或绝对路径) |
| `slides[].notes` | 否 | 演讲者备注，写入成品的备注页 |

成品页顺序 = `slides` 数组顺序。

## 槽位 ID 规则

槽位 ID 来自 `inspect_template.py` 的 manifest，全部是字符串：

- 文本槽 / 图片槽：形状 ID，如 `"14"`（manifest `text_slots[].id` / `image_slots[].id`）
- 表格单元格：`"<形状ID>!r<行>c<列>"`，1-based，如 `"9!r2c3"`（表格在 manifest `table_slots`，`cells` 字段给出每格样本）

## manifest 中与填写直接相关的字段

- `text_slots[].sample` / `sample_len`：模板演示文案及长度（不含换行符），是最可靠的字数参照
- `text_slots[].max_chars` / `max_lines`：估算预算，超出会在 `--check` 时告警
- `text_slots[].paragraphs`：样本分段数，新文案分段数尽量一致
- `text_slots[].decorative`：`true` 的槽不要填
- `text_slots[].layout_prompt`：空占位符的用途提示（取自版式，如"标题"）
- `text_slots[].group_path`：该文本框嵌在哪个组合里，仅帮助定位
- `image_slots[].aspect`：槽位宽高比。替换图与其相差过大时，引擎居中裁切，构图重要的图请预先裁好
- `table_slots[].rows` / `cols`：行列数固定，不能增删

## 校验行为

`build_deck.py plan.json --check`：

- **E（阻断）**：source 页码越界、槽位 ID 不存在、表格行列越界、图片文件不存在
- **W（放行但需复核）**：文本超预算 15% 以上、单行槽填多段、所选页面有未填的非装饰文本槽（= 模板文案将残留）、图片格式少见

不带 `--check` 直接生成时同样先跑校验，有 E 拒绝生成。
