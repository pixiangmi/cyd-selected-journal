# 更新记录

## 0.2.0 — 2026-10-05

- 新增 `--field-names original|zh`，选择原始英文或中文导出字段名。
- 新增 `--drop-empty-columns`／`--no-drop-empty-columns`，按表删除或保留整列为空的字段，空表保留表头。
- 新增 `--missing-report none|terminal|file|both`，输出各表逐列缺失数、缺失率、删除标记和可确认的缺失原因。
- JSON 保留完整原始结构，非默认表格导出附带 `export_view`；缺失统计包含已删除列。
- 导出选项保存在任务中，恢复时沿用，也可覆盖；中断时继续保存当前结果及统计。
- 文档归入 `docs/`，新增中文导出样例；本地 Cookie 归入受 Git 忽略的 `.local/credentials/`。
- 离线测试扩展至 71 项，覆盖导出组合、空表、缺失值定义、配置校验和恢复／中断。

## 0.1.0 — 2026-10-05

- 独立 LetPub 名称／ISSN 批量查询 CLI，精确匹配、去重及 Excel／CSV／JSON 导出。
- 请求间隔、有限重试、详情缓存、可选 Cookie、逐项检查点和断点恢复。
- 33 项离线测试，5 本期刊的匿名与登录字段在线验证。
