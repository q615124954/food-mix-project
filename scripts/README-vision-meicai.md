# 美菜电子小票视觉模型对照

实验脚本，只用来人工对比，不写回复核页。

复核页里，美菜电子小票默认走 **Qwen3.5-Flash**（`POST /v1/chat/completions`，见 `meicai_vision.py`）。要改用 Qwen3-VL-Plus，设置 `MEICAI_VISION_MODEL=Qwen3-VL-Plus`。密钥仍是 `HUNYUAN_API_KEY`（空则 `MOLLI_ARK_API_KEY`）。手写采购单的生产默认仍是 **HunyuanOCR**（`POST /v1/images/ocr`）。

Shot3 煲汤袋实收（模型常见 3/14.2，金标 1/4.8）是已知缺口，默认模型不为此改成更贵的型号。

没有密钥时脚本照样打印对照表，但**不发请求**。行命中、实付一致、编造行、耗时留空，这些空格子不是实测 0。仓库里不预写实跑数字。运维配好密钥后重跑，表才会填上。

## 密钥

模力方舟访问令牌。两个名字同一把钥匙，只填一个：

- 先读 `HUNYUAN_API_KEY`
- 为空再读 `MOLLI_ARK_API_KEY`

不要把钥匙写进仓库或命令行参数。

## 样本

仓库里已有同类金标图片，这三张 Shot 放在 `samples/vision_exp/`。

| 图片 | 金标 | 说明 |
| --- | --- | --- |
| `samples/vision_exp/shot1.jpg` | `samples/gold/meicai_shot1.json` | 复核后的落地行（`日物料采购_复核后期望行`） |
| `samples/vision_exp/shot2.jpg` | `samples/gold/meicai_shot2.json` | 同上 |
| `samples/vision_exp/shot3.jpg` | `samples/vision_exp/fixture_shot3.json` | 整图 OCR 解析夹具：货品在，实收数量和金额是空的。这不是已填实收的复核金标。有实收金标后换成那个路径 |

对照按你传入的 JSON 计算，脚本不改金标。

## 怎么跑

默认视觉模型是 `Qwen3.5-OCR` 和 `Qwen3-VL-Plus`。基线是现有 HunyuanOCR 加 `pipeline_min.parse_meicai_text`（窄长图仍按复核页：原宽分条，整图「收」多出至少 3 行才改用整图）。

```bash
python3 scripts/vision_meicai_compare.py \
  samples/vision_exp/shot1.jpg samples/gold/meicai_shot1.json \
  samples/vision_exp/shot2.jpg samples/gold/meicai_shot2.json \
  samples/vision_exp/shot3.jpg samples/vision_exp/fixture_shot3.json \
  --json samples/vision_exp/out/compare.json \
  --md samples/vision_exp/out/compare.md
```

不写参数时用上面这三对。`--dry-run` 有密钥也不发请求，用来看表的形状：

```bash
python3 scripts/vision_meicai_compare.py --dry-run
```

追加别的视觉模型（默认两个仍在）：

```bash
python3 scripts/vision_meicai_compare.py --model Qwen3-VL-32B-Instruct --dry-run
```

## 视觉接口

`POST https://ai.gitee.com/v1/chat/completions`。用户消息里一张图加一段字：

- `{"type":"image_url","image_url":{"url":"data:image/jpeg;base64,..."}}`
- `{"type":"text","text":...}`

提示词要求只输出 JSON，货品行字段与复核页一致：`品名`、`数量`、`金额`、`单位`、`类别`、`渠道`、`订单号`、`实收缺失`、`票面单价`。脚本再补 `单价`、`勾稽`，并按品名+单位合并，和复核页看到的一行一致。基线品名仍走现有解析（含已有别名表）。视觉侧的品名按模型原文比较，不再套那张 OCR 别名表。

## 坐标框

`Qwen3.5-OCR` 若只回 bbox / polygon 这类坐标框，状态写 **接口形态不合格**。不把框里的字折成命中数，避免读成「认了但全是 0」。

默认再请求一次回退模型 **Qwen3.5-Flash**（`--fallback-model` 可改名，`--no-fallback` 关掉）。只有模型名里带 `OCR` 且结果是坐标框时才回退。回退行平时在表里，状态是 `not_run` 或无密钥时的 `skipped_no_api_key`。

## 表里这几列

| 列 | 含义 |
| --- | --- |
| 行命中 | 金标行里，品名相同、单位不冲突、数量和金额到分都一致的行数。一对一，不重复计算 |
| 金标行 | 金标文件里的行数。这列在无密钥时也有，因为它不是模型输出 |
| 实付一致 | 模型 `paid` 与金标 `paid` 到分相同为「是」。没发请求时留空。模型交了 JSON 但 `paid` 为空、金标有实付，记为「否」 |
| 编造行 | 模型多出来、对不上金标的货品行 |
| 耗时ms | 该次 HTTP 耗时。没发请求时留空 |

JSON 与 Markdown 同一份结构，`schema` 为 `vision_meicai_compare.v1`。`contains_live_measurements` 为 false 时，文件里没有在线数字。
