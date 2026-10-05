# Cookie 登录字段测试

日期：2026-10-05（Asia/Shanghai）。本次使用用户提供的 Cookie，独立于先前匿名验收结果。

## 登录与运行结果

Cookie 已转换为 Netscape 格式，现归入 `.local/credentials/letpub.cookies.txt`，权限为 0600；`git check-ignore` 确认文件受 Git 忽略。载入 10 个 Cookie，仅发送到 LetPub 域，未保存账号密码。Cookie 值未写入本报告。以下命令中的路径已按整理后的目录更新。

运行：

```sh
.venv/bin/python -m selected_journal query \
  --input examples/acceptance.csv \
  --output runs/cookie-test-20261005 \
  --cookie-file .local/credentials/letpub.cookies.txt --refresh
```

9 条输入最终得到 7 条匹配，按期刊 ID 去重后为 5 本；另有刻意构造的 1 条未找到和 1 条 ISSN 校验失败。五本期刊的 `impact_factor.status` 和 JCR `status` 全部为 `ok`，此前匿名访问时这些字段均为 `login_required`。

PLoS One 在第一轮遇到一次 HTTP 请求失败；使用 `--resume` 单项续跑后成功，其他完成记录没有重复抓取。

## 实际读取的数据

以下为 LetPub 详情页所展示的 IF，页面注明“数据来源于网友提供”，版本标签为“2025-2026最新IF”。保留页面原值，不据此推断某一指标年份或声称已经和官方 JCR 数据核对。

| 期刊 | LetPub ID | 页面 IF | JCR 按 JIF 分区 | JCR 按 JCI 分区 |
| --- | --- | --- | --- | --- |
| PROTEIN AND PEPTIDE LETTERS | 6969 | 2.496 | Q3 | Q4 |
| NATURE | 6054 | 56.099 | Q1 | Q1 |
| SCIENCE | 7393 | 47.299 | Q1 | Q1 |
| CELL | 1562 | 45.103 | 两个学科均 Q1 | 两个学科均 Q1 |
| PLoS One | 6775 | 2.803 | Q2 | Q2 |

JCR 页面版本为“2025-2026年最新版”；JIF、JCI 分开保存，Cell 的两个学科都保留。没有把页面总览的“WOS分区等级”替代详细的 JIF/JCI 表格。

额外重新读取 ID 6969 的登录后字段 HTML：IF 单元格可见值为 2.496，JIF 表格为 Q3，JCI 表格为 Q4，与输出一致；这些字段中的结果不是隐藏干扰值。

## 输出检查

结果位于 `runs/cookie-test-20261005/`：

- `results.xlsx` 及四个同名 CSV：期刊汇总 5 行、分区明细 34 行、查询记录 9 行、候选清单 0 行。
- `results.json`：完整字段状态、版本标签及输入关联。
- `checkpoint.jsonl`：保留第一轮失败与续跑成功的记录。

已比较 Excel 和 CSV 的所有表头、行数、单元格值，并检查与 JSON 投影一致。已检查结果文本与 Excel 不含会话 Cookie 值；输入快照和检查点也不含 Cookie 值。登录详情缓存与匿名缓存分开。

退出码为 2，是因为验收名单中包含未找到和无效输入；五本期刊的登录字段读取均成功。此验证证明本次 Cookie 可用于当前页面，后续会话是否仍有效由服务器返回决定。
