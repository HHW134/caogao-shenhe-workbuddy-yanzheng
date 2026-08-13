# 草案审核 WorkBuddy 验证

标准草案（.docx）自动审核系统：上传文档 → 按可配置规则检查 → 输出问题清单 + 把每条问题以 **Word 批注**写回文档。

本仓库把此前分散在五个工作区的方案整合为一套引擎、一套规则库、一个前端：

| 原工作区 | 沉淀内容 | 本仓库位置 |
| --- | --- | --- |
| 术语代码规则梳理 | 术语和定义章节 17 条规则 | `rules/terminology_rules.json`、`tools/generate_terminology_rules.py` |
| 术语章节审核代码生成 | 术语章节审核器、引用文件审核器、双章节联动主控 | `review_engine/terminology.py`、`references.py`、`pipeline.py` |
| 规范性引用文件规则梳理 | A4R001~A4R015、FVT001~FVT008 规则 | `rules/a4guifan_rules.json`、`rules/form_validation_term_ref_rules.json` |
| 规范性引用文件审核代码 | 规则驱动的引用章节审核 CLI | `review_engine/normative_refs.py` |
| 文档审核 | 全文档结构审核（156 条规则明细）+ 批注器 | `review_engine/fulldoc.py`、`rules/fulldoc_rules.json` |

---

## 一、能力总览

三个审核模块可任意组合，结果汇总到同一张问题表：

| 模块标识 | 名称 | 覆盖内容 | 规则来源 |
| --- | --- | --- | --- |
| `chapters` | 术语与引用联动审核 | 术语条目双语/大小写/引导语/条款约束/正文使用；引用条目编号、层级、年份、占位符；**两章节联动**（术语「来源」标准不得列入规范性引用文件） | `terminology_rules.json` + `form_validation_term_ref_rules.json` |
| `normative` | 规范性引用文件规则审核 | A4R001~A4R015，完全由 JSON 驱动，批注文案与级别改 JSON 即生效 | `a4guifan_rules.json` |
| `fulldoc` | 全文档结构审核 | 主程序、目次、范围、规范性引用文件、术语、缩略语、正文 | `fulldoc_rules.json`（156 条明细，按「启用状态」执行） |

`chapters` 与 `normative` 都覆盖引用章节，但实现思路不同（启发式解析 vs 规则表驱动），互为交叉验证；同时勾选可发现单一引擎漏判的问题。

### 统一问题模型

所有模块的输出都归一化为同一条记录：

```
序号 | 模块 | 审核要点 | 级别(error/warning/info) | 问题 | 位置 | 修改建议 | 段落索引 | 原文
```

### 产物

- 带 Word 批注的 docx：每条问题锚定到对应段落，批注含「审核要点 / 级别 / 问题 / 建议」
- 问题清单 xlsx：`概览` + `审核问题` 两个工作表，按级别着色、冻结表头、自动筛选
- JSON 结果：便于二次处理与回归比对

---

## 二、快速开始

```bash
# 1. 安装依赖（建议虚拟环境）
pip install -r requirements.txt

# 2. 生成合成样例（不含真实草案）
python tools/generate_samples.py

# 3a. 命令行审核
python cli.py samples/sample1_violations.docx --all -o output

# 3b. 或启动网页版
python server.py --port 8000 --open
```

浏览器打开 <http://127.0.0.1:8000> 即可拖拽上传、勾选模块、查看问题表并下载产物。

### 命令行

```bash
python cli.py <docx> [-m chapters,normative,fulldoc] [--all]
              [-o 输出目录] [--json 结果.json] [--no-docx] [--no-xlsx]
```

终端直接打印问题表格与修改建议汇总：

```
+-------+--------------------------+--------+--------------------------------------+
| 序号  | 审核要点                 | 级别   | 问题                                 |
+-------+--------------------------+--------+--------------------------------------+
| 11    | 规范性引用文件-来源联动  | 错误   | 该标准出现在术语和定义章节中，不应放在规范性引用文件中，应放在参考文献章节中  |
+-------+--------------------------+--------+--------------------------------------+
注：所有模块的 docx 批注正文均不再展示「【规则名】[级别]」「[规则编号] 规则名」
「【审核要点】」这类规则标签前缀，仅保留「问题 / 建议 / 详情」正文（见 terminology.py
的 _build_comment_lines()、normative_refs.py 的 apply_comments()、fulldoc.py 的 _add()）。
```

### Python API

```python
from review_engine import run_review

summary = run_review("草案.docx", modules=["chapters", "normative"], out_dir="output")
print(summary["问题总数"], summary["级别统计"])
for issue in summary["问题列表"]:
    print(issue["序号"], issue["审核要点"], issue["问题"], issue["位置"])
```

### Web 接口

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| GET | `/api/modules` | 模块清单与规则库概览 |
| GET | `/api/rules/<key>` | 某规则文件的规则明细（terminology / a4guifan / linkage / fulldoc） |
| POST | `/api/review` | multipart 上传 docx（字段 `file`、`modules`），返回统一问题表 |
| GET | `/api/runs` | 历史审核记录 |
| GET | `/api/download/<run_id>/<name>` | 下载产物 |

服务仅使用 Python 标准库，无需 Flask/FastAPI。

---

## 三、规则配置

规则全部外置在 `rules/`，**改 JSON 后重跑即可生效，无需改代码**：

```
rules/
├── terminology_rules.json                 术语和定义 17 条
├── a4guifan_rules.json                    A4R001~A4R015
├── form_validation_term_ref_rules.json    FVT001~FVT008（两章节联动）
└── fulldoc_rules.json                     全文档 156 条明细（含启用状态）
```

- `a4guifan_rules.json` / `fulldoc_rules.json` 里的 `comment_text`、`how_to_fix`、`severity`、`启用状态` 直接决定批注文案、修改建议、级别与是否执行。
- `terminology_rules.json`、`form_validation_term_ref_rules.json` 目前作为规则说明与前端「规则中心」的数据源，判定逻辑在 `review_engine/terminology.py`、`references.py` 中实现。
- 网页版「规则中心」页可检索全部规则，确认某条规则是否启用、批注文案是什么。

规则说明文档见 `docs/`。

---

## 四、目录结构

```
├── cli.py                      命令行入口
├── server.py                   本地 Web 服务（标准库实现）
├── review_engine/
│   ├── __init__.py             统一入口 run_review()：模块编排 + 结果归一化 + xlsx 汇总
│   ├── pipeline.py             双章节审核主控（术语 → 引用 → 联动）
│   ├── terminology.py          术语章节解析与审核 + 批注写入 + xlsx 导出
│   ├── references.py           规范性引用文件章节解析与审核（启发式）
│   ├── normative_refs.py       规范性引用文件规则驱动审核（A4R001~A4R015）
│   └── fulldoc.py              全文档结构审核 + 批注器
├── rules/                      规则库（JSON，可热改）
├── frontend/                   前端页面（原生 HTML/CSS/JS）
├── tools/                      样例生成、规则表导出脚本
├── docs/                       规则说明与使用文档
├── samples/                    样例文档（真实草案不入库）
├── runs/                       Web 服务运行产物（不入库）
└── output/                     CLI 产物（不入库）
```

---

## 五、已知限制

1. **A4R010 / 正文引用检查是启发式**：正文需出现「应符合 / 应遵循 / 按照 / 依据 / 应采用」等规范性措辞才判定为已引用，可能产生偏多 warning。
2. **`chapters` 与 `normative` 对同一文档的引用条目解析结果可能不同**：前者依赖段落结构，遇到表格化或异形排版会解析为 0 条，此时以 `normative` 结果为准。
3. **`fulldoc` 中前言、附录模块在规则 JSON 中标记为未启用**，默认跳过。
4. 批注锚定依赖段落索引，章节级问题（如「章节缺失」）无法锚定到具体段落，只进入 xlsx 清单。
5. 解析基于段落文本与正则，建议结合人工复核。

## 六、安全提示

仓库为公开仓库，`.gitignore` 已排除 `samples/*.docx`、`runs/`、`output/`。请勿提交未发布的真实标准草案与审核结果。
