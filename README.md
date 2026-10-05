# LetPub 期刊批量查询工具

基于 [nickchen121/cyd-selected-journal](https://github.com/nickchen121/cyd-selected-journal) 的 fork，新增独立 Python 命令行工具：输入期刊名称或 ISSN，批量查询 LetPub 并导出 Excel、CSV 和 JSON。

- Python 3.11+，无需 ChromeDriver 或 Tkinter。
- 支持 TXT、CSV、XLSX 和标准输入名单，精确匹配名称／ISSN，按期刊去重。
- 获取页面 IF、CiteScore、OA、审稿周期、JCR 的 JIF／JCI 分区、期刊分区表和新锐分区，并保留版本与来源。
- 支持可选本地 Cookie、24 小时缓存、断点续跑和有限重试。

## 快速开始

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -e '.[test]'
.venv/bin/python -m selected_journal query --input examples/journals.csv
```

每次任务会在 `runs/` 下生成结果。需要登录才能查看的字段，可通过 `--cookie-file` 读取本地 Netscape Cookie 文件；Cookie、缓存和运行目录均受 Git 忽略。

详细说明见 [命令行使用指南](README_CLI.md)、[验证记录](VALIDATION.md) 和 [登录字段测试记录](COOKIE_TEST.md)。仓库提供 [输入模板](examples/journals.csv) 及 [匿名查询样例输出](examples/output/)。

```sh
.venv/bin/python -m pytest -q
```

当前有 33 项离线测试，以及 5 本期刊的匿名／登录字段在线验证。输出是 LetPub 页面抓取记录；指标年份、分区体系和缺失状态分别保留。

## 上游说明

原始脚本 `cyd_v1_stable.py` 和 LICENSE 保留。以下为上游原 README 内容：

[TOC]

B站，水论文的程序猿，https://space.bilibili.com/383551518
独立开发，开源代码供粉丝、同学们研究，禁止商用，盗版侵权必究

# 百度学术代码思路

1. 用pyinstaller打包成exe程序
2. 用tkinter设计可视化界面
3. 用selenium通过==关键词（这是一个可以有很大的改进的地方）==获取百度学术论文
   1. 获取百度学术论文后通过requests获取学术论文的期刊
4. 用request请求letpub获取期刊级别
5. 通过逻辑语句对期刊划分为3种
   1. 中文期刊
   2. 会议
   3. 英文会议
6. 期间获取一篇期刊保存一篇期刊
   1. 保存时会按照特定规则排序

## 改进点()

1. 通过关键词搜索，可以增加更多的百度学术高级搜索方式(2023年4月26日已实现)；
2. 中文期刊可以通过知网期刊搜索，增加是否为核心期刊；
3. letpub无法获取期刊IF，可以增加IF排序；
4. 增加谷歌学术等平台的论文搜索功能；
5. ……
