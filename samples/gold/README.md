# 8/11–8/13 金标样本

这是 2026-08-11、08-12、08-13 三天的公开金标，用来对照采购单识别结果。对比时以 `material_gold.xlsx` 为准。

## 文件

| 文件 | 内容 |
|------|------|
| `receipt_2026-08-11.jpg` | 8/11 采购单照片 |
| `receipt_2026-08-12.jpg` | 8/12 采购单照片 |
| `receipt_2026-08-13.jpg` | 8/13 采购单照片 |
| `material_gold.xlsx` | 金标物料表（公司物料采购、肉类采购、日物料采购、员工餐、综合表） |
| `baseline_aug11_13.json` | 从金标表抽出的这三天明细（肉类采购、日物料采购、员工餐） |
| `baseline_overrides.json` | 相对原始表格的人工修订 |
| `summarize_gold.py` | 只读摘要：统计这三天有值的格子，确认样本能读开 |

8/13「鸡毛菜」的数量是 **5.8**，金额是 **28**。原始表格里数量曾误写成 58，金标表已改成 5.8。

## `baseline_overrides.json`

overrides 只记录需要改掉的笔误，不替代整张表。当前只有一条：

- 表：日物料采购
- 日期：2026-08-13
- 品名：鸡毛菜
- 数量：5.8（覆盖误写的 58）
- 金额：28（不改）

`material_gold.xlsx` 和 `baseline_aug11_13.json` 里这条已经是 5.8 / 28。对账时以 xlsx 的数量和金额为准。表内单价列经常是空的；JSON 里的单价是金额除以数量的派生值。

## 查看摘要

不调用 OCR，也不需要 API 密钥：

```bash
python3 samples/gold/summarize_gold.py
```

## 密钥

本目录不含 OCR API 密钥。百度 OCR 密钥只从环境变量 `BAIDU_OCR_API_KEY` 读取，不要写进仓库，也不要提交 `.env`。
