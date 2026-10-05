# LetPub 期刊批量查询工具

输入期刊全称或 ISSN，批量读取 LetPub 当前信息，按期刊去重并导出 Excel、CSV、JSON。适合每批 100 本以内；无需 Chrome、ChromeDriver 或 Tkinter。

上游 `cyd_v1_stable.py` 保留原样，原 README 内容保留在首页“上游说明”中。新工具使用独立的 `selected_journal` 包，不启动旧脚本，不搜索论文。

## 开始使用

本次开发已在仓库内创建 `.venv`。在终端中运行：

```sh
cd '/Volumes/Mac拓展/tools/selected_journal/cyd-selected-journal'
.venv/bin/python -m selected_journal query --input examples/journals.csv
```

其他机器需 Python 3.11 或更新版本，先安装：

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -e '.[test]'
```

`pyproject.toml` 声明运行依赖与测试依赖，`requirements-dev.lock` 固定本次验证过的完整开发依赖版本。需要复用这些版本时先运行 `pip install -r requirements-dev.lock`。

每次新任务使用新目录，默认位于 `runs/日期-时间/`。指定目录示例：

```sh
.venv/bin/python -m selected_journal query \
  --input examples/journals.csv --output runs/my-journals
```

## 输入名单

支持 TXT、CSV、XLSX 和粘贴输入。TXT 或粘贴每行一个期刊全称或 ISSN，不以逗号拆分名称。空行忽略，ISSN 支持无连字符和末位 X；错误校验位保留为无效记录。

```sh
.venv/bin/python -m selected_journal query --stdin
```

粘贴名单后按 Ctrl+D 结束输入。也可通过标准输入管道传入 TXT：

```sh
cat examples/journals.txt | .venv/bin/python -m selected_journal query --stdin
```

CSV 和 XLSX 使用下面的列名，至少存在一列。XLSX 读取第一个工作表，ISSN 单元格应使用文本格式，避免丢失前导零。示例文件位于 `examples/`。

```csv
journal_name,issn
PROTEIN AND PEPTIDE LETTERS,
,0929-8665
PROTEIN AND PEPTIDE LETTERS,1875-5305
```

有 ISSN 时优先查询 ISSN，匹配详情页的 ISSN 或 eISSN；同时填写名称时，全称也必须一致。仅填名称时采用忽略大小写、规范空白后的完整名称精确匹配，不自动选择相似期刊。

如果得到多个候选、名称不完整、名称与 ISSN 冲突，或搜索超过 10 页仍未结束，记录为待确认；候选清单提供名称、ISSN 和链接。核对后补充 ISSN，作为新名单重新运行。查询期刊缩写通常需要这个步骤。

## 可选登录与缓存

默认匿名访问。LetPub 部分字段需要登录；新工具只保留可访问的数据，不会从网页标题、搜索摘要、实时 IF 或隐藏文字补填正式 IF。

如需登录字段，可在浏览器登录 LetPub 后，将该站点的 Cookie 导出为 **Netscape HTTP Cookie File** 格式，保存在本地，用如下参数读取：

```sh
.venv/bin/python -m selected_journal query \
  --input examples/journals.csv --cookie-file /absolute/path/letpub.cookies.txt
```

不接收账号密码。只载入 LetPub 域的未过期 Cookie，日志和结果不包含 Cookie 值。是否实际登录成功由返回页面决定；Cookie 过期或字段仍受限时，公开信息照常保存，受限字段标记 `login_required`。真实登录路径已使用用户提供的 Cookie 在线验证 5 本期刊，记录见 `COOKIE_TEST.md`。

详情缓存有效期 24 小时，存放于任务目录的父目录下 `.letpub-cache/`。匿名与登录缓存分开，不同 Cookie 集合也隔离。加 `--refresh` 跳过磁盘缓存，同一批次内仍避免重复读取同一期刊。

## 进度、中断与恢复

顺序访问，请求启动之间至少间隔 3 秒。连接／读取超时分别为 10／30 秒。瞬时错误最多尝试 3 次，重试等待 5／15 秒。遇验证码、HTTP 403 或 429 时停止本批网络访问，保存当前结果。

Ctrl+C 会保存已完成结果。恢复原任务：

```sh
.venv/bin/python -m selected_journal query --resume runs/my-journals
```

恢复时读取保存的输入快照，跳过已匹配、未找到、待确认和无效记录，重试失败／阻断／部分解析失败记录，以及尚未处理的输入。不与 `--input`、`--stdin` 或 `--output` 同用。仍可指定 Cookie 文件和 `--refresh`；已经完成的记录不会因此重新查询。

新 Cookie 或希望更新所有期刊时，使用原名单启动新任务；若已知公开字段解析发生变化，可配合 `--refresh`。

## 结果文件与含义

每个任务生成：

- `inputs.json`：输入快照，不依赖原文件继续存在。
- `checkpoint.jsonl`：逐项追加的检查点；恢复时只修复末尾未写完的记录。
- `results.xlsx`：期刊汇总、分区明细、查询记录、候选清单四张表，含筛选和冻结表头。
- 四个同名 CSV：UTF-8 BOM 编码，便于 Excel 打开。
- `results.json`：完整字段、状态、分区数组和输入关联，`schema_version=1`。

汇总以 LetPub 期刊 ID 去重，保持输入首次出现顺序；`input_ids` 可以追溯同一期刊的多条原始输入。默认使用原始英文列名，工作表名称为中文；可选择中文列名及删除空列，见下节。

主要字段是 `journal_name`、`abbreviation`、`issn`、`eissn`、`impact_factor`、`citescore`、`oa`、`review_time`，另有 `source_url`、`fetched_at`。每个字段有 `_status`、`_raw`、`_version`、`_year` 列。抓取时间为带时区的 ISO 时间。

- `ok`：已解析。
- `not_provided`：页面未提供或明确没有记录。
- `login_required`：需要登录才能查看。
- `parse_error`：发现对应字段但无法可靠解析。

分区按 `jcr`、`journal_partition`（页面“期刊分区表”）、`xinrui` 分别保存最新版本，保留来源名称；不会把新锐与旧版表合并或替换名称。每个学科单独一行，JIF、JCI 等依据保留在 `basis` 中。页面只有发布年份时，不推断 IF／CiteScore 的指标年份，年份为空但版本标签保留。JCR 的年份范围也不自动选其中一年。

`oa` 使用 `open`、`hybrid`、`closed` 或 `unknown`。审稿周期保留“网友分享经验”等原文。空值不能当作 0，未找到不能当作“未收录”。

查询状态包括 `matched`、`needs_confirmation`、`not_found`、`invalid`、`failed`、`partial`、`blocked`、`pending`。字段因登录受限仍可算匹配成功；解析失败则记录 `partial`。

退出码：全部匹配为 0；存在未找到、待确认、无效或失败／未完成输入为 2；输入配置或保存错误为 1；Ctrl+C 为 130。

## 自定义导出与缺失统计（v0.2.0）

| 参数 | 可选值 | 默认行为 |
| --- | --- | --- |
| `--field-names` | `original` 原始英文，`zh` 中文 | `original` |
| `--drop-empty-columns` | 删除整列为空的列 | 保留全部列 |
| `--no-drop-empty-columns` | 明确保留全部列，用于覆盖恢复任务的设置 | 保留全部列 |
| `--missing-report` | `none`、`terminal`、`file`、`both` | `none`，不输出统计 |

例如，中文表头、删除空列，同时显示终端统计并保存统计文件：

```sh
.venv/bin/python -m selected_journal query \
  --input examples/journals.csv --output runs/chinese-journals \
  --field-names zh --drop-empty-columns --missing-report both
```

`--field-names` 只翻译字段名，不翻译期刊名、页面原文、状态代码或 OA 值。例如 `impact_factor` 为“影响因子IF”，`impact_factor_status` 为“影响因子IF_状态”，`journal_partition_subjects` 为“期刊分区表（中科院）_学科分区”。原始体系名称依然来自页面。

空值定义为 `null`、空白字符串、空列表或空对象；`0`、`False`、字符串 `unknown` 是有效值。删除空列按四张表各自的数据行判断，**没有数据行的表保留全部表头**。只删除整列为空的列，不删除部分缺失的列，也不删除有值的状态列。恢复后新增记录可能让先前的空列重新出现在输出中。

Excel 与 CSV 使用相同的列名和保留列。`results.json` 始终保留 `journals`、`queries`、`candidates` 的完整原始字段和状态，以便程序读取和追溯。选择中文字段或删除空列时，JSON 另含 `export_view`，其中 `columns` 记录原始字段与输出字段名的对应关系，`rows` 使用相同的中文或英文键及保留列。输入快照、检查点和缓存不因导出设置改变。

缺失统计基于**删空列之前**的四张表。期刊汇总按去重后的期刊数计算；查询记录按输入条数计算，分区明细按学科分区行数计算：

- `terminal`：终端逐表显示缺失字段的“缺失数／总行数（百分比）”，标注已删除列，并列出无缺失列数量。
- `file`：保存 `missing_values.json` 和 UTF-8 BOM 的 `missing_values.csv`，不打印终端统计。
- `both`：同时输出上述两种形式。
- `none`：不输出统计；恢复任务时会移除该目录中旧的统计文件，避免误读过期结果。

统计文件包含表名、原始字段名、输出字段名、总行数、缺失数、非缺失数、缺失率和是否删除列。`missing_rate` 是 0～1 的比例，如 `0.4` 代表 40%；空表缺失率为空值，不视作 0% 或 100%。文件列出全部字段，包括无缺失和已删除的字段。`reasons` 对核心字段和分区内容保留可关联的 `login_required`／`not_provided`／`parse_error` 计数；无法明确关联的空值计为 `unspecified`，不推测原因。指标年份无法确认时为空，也会计入统计。

导出设置保存在 `inputs.json` 的 `export_options` 中。`--resume` 默认沿用；显式参数可覆盖并保存，已完成任务可用此方式重新导出，无需重复抓取：

```sh
.venv/bin/python -m selected_journal query --resume runs/chinese-journals \
  --field-names original --no-drop-empty-columns --missing-report file
```

旧版本任务没有保存导出设置时，恢复使用默认设置。用户中断或批次受阻时，也按所选设置导出已完成数据及统计；缺失统计本身不改变退出码。

## 验证与样例

```sh
.venv/bin/python -m pytest -q
```

测试包括匿名样例页面、模拟登录后的 IF／JCR 表格、隐藏值、分页、输入校验、精确匹配、请求间隔、重试、缓存、Cookie 域过滤、检查点恢复以及 Excel／CSV／JSON 一致性。模拟登录测试与实际 Cookie 在线测试分别记录，在线结果见 `COOKIE_TEST.md`。

`examples/acceptance.csv` 是在线验收名单，含 5 本不同期刊、同刊名称／ISSN／eISSN、一个不存在的名称和一个无效 ISSN。`examples/output/` 提供本次匿名查询结果；`VALIDATION.md` 记录验证范围及结果。这些样例是抓取当时的记录，运行新任务获取当前信息。

仓库的 LICENSE 和原作者 README 中的限制文字存在不一致，此处不改变上游声明。
