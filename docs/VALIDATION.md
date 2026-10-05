# 验证记录

日期：2026-10-05（Asia/Shanghai）。上游基线提交为 `cdcd9d06dc3960ebe3b82788ad5edc944680dac2`；原脚本和 LICENSE 未修改。发布 fork 时在 README 首页新增工具介绍，原 README 内容保留在“上游说明”中。

## v0.2.0 导出功能追加验证

在同一 `.venv` 中运行 `python -m pytest -q`：**71 passed**（含下文首版 33 项测试）。未额外联网抓取；本次变更涉及导出和任务选项，使用首版已核对的匿名数据与固定 HTML 样本验证。

- 原始／中文字段 × 保留／删空列 × 不统计／终端／文件／两者共 16 组，逐项核对 Excel 与 CSV 所有表头、行数和单元格，JSON 的 `export_view` 与原始数据一致。
- 四张表的全部字段都有中文映射，每张表无同名字段冲突。
- 缺失定义验证 null、空白、空容器，以及有效的 0、False、unknown；混合记录的计数和比例正确。
- 删除空列后仍可统计原始全部字段；登录受限的 IF 记录 `login_required`，其状态列保留。
- 空表保留全部表头，缺失率为 null；汇总按去重期刊数，查询按输入数计数。
- 新任务保存导出设置，恢复继承或显式覆盖；旧任务没有导出设置时使用默认值，完成记录无需联网重导出。
- 文件／终端统计可独立选择，关闭文件统计会移除旧统计文件；非法配置返回 1。
- Ctrl+C 后按选择保存部分结果和统计，退出码保持 130；中文导出的公式文本仍保持安全的文本单元格。
- 固定匿名样例用中文、删空列和文件统计重导出到 `examples/output-options/`，原始 `examples/output/` 未修改。
- 本地保存的匿名、登录任务各 5 本期刊／9 条输入实际恢复重导出；拦截所有网络请求，确认请求数为 0。四表内容、原始 JSON 数据和缺失统计逐项一致，退出码均为预期的 2（样例含未找到／无效输入）。
- 可编辑安装更新为 0.2.0，两种 CLI 入口正常；文档链接与两套样例 SHA256 全部有效，待提交文本及 XLSX 内部 XML 未包含本地 Cookie 值。

工作区文档集中在 `docs/`；本地任务和缓存保留，Cookie 移至 `.local/credentials/letpub.cookies.txt` 并保持 0600 权限，不提交 Git。

以下是 v0.1.0 初始实现与在线验证记录，保留当时数据和验证边界。

## 离线检查

在仓库 `.venv`（Python 3.13.0）中运行 `python -m pytest -q`：**33 passed**。

覆盖内容：

- 从匿名公开页面提取的固定 HTML 样本，按字段标签及结果表定位。
- 隐藏分区数字、CSS 隐藏元素、推荐期刊、文章链接及实时 IF 的排除。
- 模拟登录页面的 IF 年份、JIF/JCI 分区及多学科保存；不同年份分区按最新版本选择。
- 名称、ISSN、eISSN、名称冲突、多候选、分页上限、重复分页及分页失败。
- 粘贴／TXT／CSV／XLSX、空名单、错误列名、无效 ISSN 与损坏 XLSX。
- 请求超时／503 重试、404 不重试、三次尝试上限、三秒间隔、验证码／403／429 停止后不再发请求。
- 24 小时缓存、强制刷新、过期缓存、内存去重和部分解析失败不写成功缓存。
- Cookie 域过滤、过期 Cookie、匿名／登录缓存隔离及 Cookie 值不进入结果。
- 中断后保存、失败或部分解析失败的续跑、末尾检查点修复、已有输出目录保护。
- Excel／CSV 表格内容一致、JSON 关联一致和表格公式文本处理。

可编辑包安装成功，`python -m selected_journal` 与 `selected-journal` 两个入口均可使用。

## 在线验收

运行命令：

```sh
.venv/bin/python -m selected_journal query \
  --input examples/acceptance.csv --output runs/acceptance-20261005
```

匿名模式，未使用账号或 Cookie。9 条输入：7 条 `matched`、1 条 `not_found`、1 条 `invalid`，去重后为 **5 本期刊**。由于名单刻意包含无效／未找到输入，退出码为 2，符合定义。

| 期刊 | LetPub ID | ISSN | 输入方式 | CiteScore | OA |
| --- | --- | --- | --- | --- | --- |
| PROTEIN AND PEPTIDE LETTERS | 6969 | 0929-8665 | 名称、ISSN、eISSN，共三条输入合并 | 2.80 | No |
| NATURE | 6054 | 0028-0836 | ISSN | 77.70 | No |
| SCIENCE | 7393 | 0036-8075 | ISSN | 44.60 | No |
| CELL | 1562 | 0092-8674 | ISSN | 67.80 | No |
| PLoS One | 6775 | 1932-6203 | 名称 | 4.80 | Yes |

以上 CiteScore 页面标签均为“2026年6月最新版”，未把该发布年份推断成指标年份。数值是本次抓取记录，后续可能变化。

详情页核对来源：[PROTEIN AND PEPTIDE LETTERS](https://www.letpub.com.cn/index.php?page=journalapp&view=detail&journalid=6969)、[NATURE](https://www.letpub.com.cn/index.php?page=journalapp&view=detail&journalid=6054)、[SCIENCE](https://www.letpub.com.cn/index.php?page=journalapp&view=detail&journalid=7393)、[CELL](https://www.letpub.com.cn/index.php?page=journalapp&view=detail&journalid=1562)、[PLoS One](https://www.letpub.com.cn/index.php?page=journalapp&view=detail&journalid=6775)。

核对时重新访问了全部 5 个详情页，将页面主体中可见的名称、ISSN/eISSN、CiteScore、OA、审稿周期、新锐分区及 2025 年升级版分区与导出逐项比较，一致。Cell 的两个小类学科均被保存。两套分区体系没有合并。

五个样例的正式 IF 和 WOS/JCR 分区均提示登录后查看，因此结果保存空值和 `login_required`；未采用标题中的数字、搜索页的 IF 或实时 IF 补填。

不存在名称 `CODEX NONEXISTENT JOURNAL 20261005` 得到未找到；`0929-8666` 的 ISSN 校验失败，不发起该行的网络请求。

输出包含期刊汇总 5 行、分区明细 27 行、查询记录 9 行、候选清单 0 行。Excel 与对应 CSV 的表头、行数及全部单元格内容一致，JSON 与汇总字段、输入关联一致。验收样例副本位于 `examples/output/`。

实际执行 `query --resume runs/acceptance-20261005` 后，9 条已完成输入全部跳过，输出统计保持一致，没有重复抓取。

## 尚未验证的边界

- 初次匿名验收时未提供有效登录 Cookie。随后用户提供 Cookie，已完成 5 本期刊的真实登录字段验证；追加记录见 `COOKIE_TEST.md`。
- 在线验收是小批量、当前页面结构验证，不代表所有期刊和所有未来页面版本都已覆盖。遇无法解析的字段或页面，工具保留明确状态，不填入推测值。
- 没有开展 100 本以上的容量测试；首版面向每批 100 本以内。
