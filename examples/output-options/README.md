# 中文导出与缺失统计样例

本目录使用 `../output/results.json` 中已核对的匿名抓取记录离线重导出，未再次查询网站。设置为：

```sh
--field-names zh --drop-empty-columns --missing-report file
```

- `results.xlsx` 与四个 CSV 使用中文表头，按表删除整列为空的列；候选表无记录，仍保留完整表头。
- `results.json` 保留原始字段、状态和输入关联，额外包含所选表格视图 `export_view`。
- `missing_values.csv`／`missing_values.json` 包含全部原始列的缺失统计；`dropped` 或“是否删除”标识从表格中删除的列。
- IF 在该匿名样例中需要登录，所以“影响因子IF”列被删除，但“影响因子IF_状态”仍保留 `login_required`；统计中也保留原因。
- 缺失率使用 0～1 比例，空表为空；这份样例是历史抓取记录，不代表网站当前值。

`SHA256SUMS` 校验本目录生成的输出文件，不含本说明文件。
