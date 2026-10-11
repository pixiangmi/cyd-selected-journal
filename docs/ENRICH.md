# 文献表期刊信息回填

v0.4.0 保留已核实但尚未修复的边界问题：结果文件含无效 ISSN 时可能整体导入失败，PDF 追加可能覆盖表下内容，回填的中断退出码也待完善。具体条件和使用建议见 [已知问题](KNOWN_ISSUES.md)。

`enrich` 读取文献 XLSX 的期刊列，统计原始名称与规范去重名称，复用现有 LetPub 查询流程，按期刊身份把全部期刊汇总字段追加到同一工作表。原列、文献顺序、其他工作表、表格名称和冻结窗格保留，输出必须是新文件。

## 一次完成查询和回填

```sh
.venv/bin/python -m selected_journal enrich \
  --input /path/literature.xlsx --output /path/literature_journals_v2.xlsx \
  --journal-column 所属期刊 --key-column "Cite key" --header-row 7 \
  --cookie-file .local/credentials/letpub.cookies.txt --refresh
```

表头默认在前 50 行自动定位；默认第一张工作表，存在 `Cite key` 时自动用它作为记录标识。可使用 `--sheet`、`--header-row` 和 `--key-column` 指定。无标识列时使用 Excel 原始行号；有标识列时空值或重复会停止。期刊列不能是公式，带总计行的文献表暂不支持。

默认在输出文件旁创建“输出文件名_资料”目录，可用 `--work-dir` 改变。保存文献输入快照、期刊去重统计 CSV、查询名单、完整查询结果、逐文献映射、追加单元格值和回填报告。查询采用中文表头、保留空列及完整缺失统计。

输出文件或原文件不能被覆盖。重复执行时提供新的输出文件名；指定已有 `--work-dir` 可以从查询检查点续跑。名单变化时必须使用新的工作目录。已有完成记录遵循 `query --resume` 行为：`--refresh` 不重查已完成项；需要更新全部指标或更换 Cookie 时使用新的工作目录。

## 使用已有结果，完全离线回填

```sh
.venv/bin/python -m selected_journal enrich \
  --input /path/literature.xlsx --output /path/literature_journals_v2.xlsx \
  --results /path/initial/results.json \
  --results /path/rechecked/results.json \
  --mappings /path/verified_mappings.json
```

`--results` 可重复，指定后不访问网络。支持读取多个任务并按 LetPub ID 合并。成功查询不会被后续失败覆盖；同一 ID 使用可解析字段较完整的记录，保留来源任务。此行为用于同一轮初查/复查，不能用来强行混合不同年份的历史指标。

## 核查映射

名称仅按 Unicode NFKC、空白、大小写去重，不自动猜测相似期刊。对于长名称、缩写、标点差异和未知刊名，应先在出版商、ISSN/NLM 等来源核实，再提供 JSON：

```json
{
  "verified_aliases": {
    "Science China. Life sciences": {
      "name": "Science China Life Sciences",
      "issn": "1674-7305",
      "source": "https://www.ncbi.nlm.nih.gov/nlmcatalog/101529880"
    }
  },
  "paper_overrides": {
    "ExampleKey": {
      "record_type": "preprint",
      "standard_name": "OSF Preprints",
      "sources": ["https://doi.org/verified-preprint-doi"],
      "note": "核实为预印本，保留该版本身份。"
    }
  }
}
```

每项映射必须提供 `source` 或 `sources`。ISSN 校验位必须正确；有 ISSN 时提交 ISSN 单独查询，避免别名和 LetPub 标准刊名不一致。单篇 `paper_overrides` 优先于期刊名称映射。使用用户已核实的映射，不自动检索论文、修改 DOI 或替换预印本为正式发表版。

bioRxiv、medRxiv、arXiv、OSF Preprints 默认标记为预印本。Unknown Journal 或空刊名保留未知状态，指标不填。其他未确认记录保留工具查询状态和说明。预印本指标状态使用 `not_applicable`，未查询字段使用 `not_queried`；工具返回的原始状态不改写。

## 输出字段及边界

全部 `SUMMARY_HEADERS` 字段使用中文表头，主要指标放在前面，完整状态、页面原文、版本、指标年份和来源放在后面；另加匹配状态、说明、核查来源及各分区年份。JCR 每个学科的 JIF/JCI 依据均保留；期刊分区表、新锐分区分开保存。IF/CiteScore 保持数值，缺失保持空白，合法 0 不丢失；页面没有明确指标年份时不推断。

保留原 XLSX 的所有工作表，只在选定文献表追加列，并扩展其原有 Excel Table/筛选范围。列宽和新增区域自动换行，原列样式保持；行高可能扩大以显示新增内容。具有复杂、不受 openpyxl 支持的专有对象的工作簿，应先备份并核查对象保留情况；本命令不支持 XLSM 或自动重算公式。

退出码：全部期刊匹配且预印本正确标记为 0；存在未确认、未知、失败或部分解析项为 2（新版表和报告仍保存）；配置/保存错误为 1。登录受限指标可以保留为空且期刊匹配成功，详见字段状态。工具保留现有请求间隔、有限重试、访问验证停止和检查点机制。

## 验证

```sh
.venv/bin/python -m pytest -q
```

新增测试覆盖 Unicode/大小写去重、稳定标识、核查映射和 ISSN 校验、未知/预印本、合法零值、指标年份不推断、多学科 JIF/JCI、跨任务成功结果保留、原文件/原列/其他工作表/样式/过滤和冻结窗格、在线调度与离线回填。


HTTP 200 的限速提示页（如“您请求页面的速度过快”）同样视为访问受阻，停止当轮请求并保存进度；不继续逐项请求，不采用规避限速的措施。


在线去重查询可加 `--request-interval 20`；已有核实结果可加
`--identity-results /path/to/previous/results.json --refresh`，只复用刊名/ISSN对应身份，
重新取得指标。重复提供该参数可使用多份身份来源。若更新Cookie，应使用新work-dir
完整刷新；`--resume`仅恢复尚未完成的输入，不会刷新已完成的匹配。


### 增量PDF资源

`selected_journal sync-pdfs --input 旧文献表.xlsx --snapshot pdf_snapshot.json --output 新查询源.xlsx`
可读取原六列表或本工具完整补充的表，保留原六列值和顺序，仅在表末追加新增PDF文献。
输出是用于重新enrich的六列表；原表保留。快照含`papers`或`all_papers`数组，
每篇字段为year、cite_key、title、journal、relevance、citations_per_year、pdf。
pdf必须为true/false或✓/X/…；只纳入true/✓，数值缺失为null。
快照可从已保存的Undermind逐页元数据汇总生成，必须保证覆盖范围和PDF状态已核验。
该命令不连接Undermind，不推断PDF状态、不改写旧文献值；Cite key缺失或重复时报错。
随后使用query/enrich刷新期刊字段；`--work-dir`另存增量核查报告。

OA值保留open、closed、hybrid、transformative、unknown类别；页面明确标注转换期刊时为transformative，原文另外保留，不推断转型是否完成。
