# YaFo 兼容采购助理  
# YaFo Compatible Purchase Assistant

> A lightweight procurement automation assistant for compatible consumables business.  
> 面向兼容耗材采购业务的轻量级自动化工具，用于把高频、重复、易出错的 Excel / XML / 成本 / 预测处理流程集中到一个 GUI 工具中。

---

## Overview｜项目简介

**YaFo 兼容采购助理** 是一个基于 Python + Tkinter 开发的采购业务辅助工具，面向兼容耗材采购场景，主要用于处理工厂直发订单、分仓订单、NC 存货资料转换、存货成本回填、新品定价、销量预测与备货建议等工作。

The tool is designed to standardize repetitive procurement workflows, reduce manual Excel operations, and improve the accuracy and consistency of order XML generation, inventory cost filling, pricing support, and replenishment analysis.

当前版本：**v3.5.39**

---

## Why This Tool Matters｜提效价值

在没有生成器之前，采购订单处理主要依赖手工 Excel 模板、复制粘贴、人工匹配 NC 存货信息、手动整理供应商文件、再生成或录入系统所需数据。该流程存在几个典型问题：

- **重复操作多**：供应商文件格式不同，需要逐个打开、清洗、复制、核对。
- **人工匹配易错**：存货编码、型号、条码、产品名称之间经常需要人工判断。
- **格式要求严格**：NC 导入 XML 对字段、节点、日期、仓库、供应商名称等要求固定，人工处理容易出错。
- **数据链路分散**：订单生成、成本回填、定价、销量预测原本是多个独立流程，难以形成统一入口。
- **追溯困难**：手工处理很难稳定保留匹配结果、异常记录、未匹配明细和处理日志。

YaFo 兼容采购助理将上述流程整合为一个桌面 GUI 工具，将“古法炮制”的 Excel 手动流程升级为标准化、批量化、可追溯的自动化处理流程。

The core value is not only file generation. It is the conversion of fragmented procurement operations into repeatable, structured, and auditable workflows.

---

## Development Background｜研发背景

本工具最初来源于兼容采购日常工作中的实际需求：供应商文件格式多、订单导入规则复杂、NC 存货匹配容易出错、工厂直发和分仓订单处理流程不统一。

早期阶段，工具分别以独立脚本形式存在：

- **分仓订单导入工具阶段**：支持北京、沈阳、南京等分仓订单 XML 生成。
- **工厂直发订单导入工具阶段**：支持多供应商文件上传、NC 产品匹配、工厂直发库 XML 生成和汇总 Excel 输出。
- **主生成器合并阶段**：将工厂直发和分仓订单整合为统一入口。
- **业务扩展阶段**：逐步集成供应商直发表格处理、新品定价、存货成本回填、销量预测与备货建议。
- **现代化阶段**：优化 GUI 界面、功能选择布局、维护日志、旧版 Excel 宏模板下载入口。

The project was developed from real procurement scenarios, not from a generic demo requirement. Its structure reflects practical business pain points: supplier format diversity, NC system constraints, inventory data matching, and the need for repeatable procurement execution.

---

## Key Features｜主要功能

### 1. 工厂直发订单 XML 生成  
### Factory Direct Order XML Generation

- 支持一次性上传多个供应商文件。
- 自动识别供应商和文件格式。
- 自动匹配 NC 存货名称。
- 批量生成 NC 可导入的采购订单 XML。
- 同步生成汇总 Excel，便于核对单据号、规格型号、收件人、快递单号等信息。
- 支持周结预定单统计表按供应商拆分生成 XML。
- 内置“古法采购订单模板”下载入口，可导出 `.xlsm` 模板并调用系统默认 Excel 打开。

### 2. 分仓订单 XML 生成  
### Warehouse Order XML Generation

- 支持北京、沈阳、南京等分仓订单。
- 可一次性上传多个分仓文件。
- 根据文件名自动识别分仓。
- 支持不同供应商分仓格式识别。
- 自动匹配 NC 存货信息并生成分仓采购订单 XML。

### 3. 供应商直发表格处理  
### Supplier Direct-shipment Table Processing

- 用于预处理预定单。
- 按供应商拆分并导出 Excel。
- 适合工厂直发前置整理场景。

### 4. 新品上新定价  
### New Product Pricing Assistant

- 集成新品定价子模块。
- 支持按品牌、成本、批发毛利率、零售毛利率进行定价测算。
- 用于新品上新前的批发价和零售价辅助判断。

### 5. 存货成本回填  
### Inventory Cost Filling

- 支持选择 NC 成本表。
- 批量导入多个供应商成本来源表。
- 支持编码优先、条码兜底的匹配逻辑。
- 回填结果先暂存在程序内，最后统一导出 Excel。
- 输出匹配明细、未匹配明细、异常明细、来源文件统计等结果。

### 6. 销量预测与备货建议  
### Sales Forecasting and Replenishment Recommendation

- 支持导入销售订单与库存信息。
- 基于历史销售数据预测未来需求。
- 支持 SKU 分层：A+ / A / B / C / D / E / 零销量。
- 支持工作日效应、趋势因子、同比修正、季节因子。
- 支持产品生命周期识别。
- 支持客户集中度风险评估。
- 输出预测销量、建议备货量和 Excel 报告。

### 7. NC 存货 CSV 导入  
### NC Inventory CSV Import

- 将 NC 导出的 CSV 转换为主生成器可读取的 NC 产品信息 Excel。
- 自动识别常见编码格式。
- 导入后自动加载为当前 NC 产品信息文件。

### 8. 维护日志  
### Built-in Changelog

- GUI 内置维护日志入口。
- 记录从早期独立工具到当前集成版本的主要迭代。
- 方便后续维护、交接和版本追溯。

---

## Workflow｜典型使用流程

### 工厂直发订单

1. 打开主程序。
2. 选择 **工厂直发订单**。
3. 选择 NC 产品信息文件。
4. 批量选择供应商订单文件。
5. 设置输出目录。
6. 点击 **生成 XML 和汇总表**。
7. 到输出目录查看 XML 文件和汇总 Excel。

### 分仓订单

1. 选择 **分仓订单**。
2. 选择多个分仓文件。
3. 程序自动识别北京 / 沈阳 / 南京等分仓。
4. 点击生成分仓 XML。

### 存货成本回填

1. 选择 **存货成本回填**。
2. 选择 NC 成本表。
3. 批量添加供应商成本来源表。
4. 点击 **开始回填 / 加入暂存**。
5. 检查结果后点击 **导出结果**。

### 销量预测与备货建议

1. 选择 **销量预测与备货建议**。
2. 添加销售订单文件。
3. 可选导入库存信息文件。
4. 设置目标周转天数和到货周期。
5. 点击分析并导出预测报告。

---

## Requirements｜运行环境

Recommended environment:

- Windows 10 / Windows 11
- Python 3.9+
- Microsoft Excel installed, recommended for `.xlsm` template and legacy Excel compatibility

Python packages:

```txt
pandas
numpy
openpyxl
chardet
statsmodels
xlrd
pywin32
```

Install dependencies:

```bash
pip install pandas numpy openpyxl chardet statsmodels xlrd pywin32
```

Notes:

- `statsmodels` is used for Holt-Winters forecasting. If unavailable, the forecasting module can fall back to other methods.
- `xlrd` is used for reading legacy `.xls` files.
- `pywin32` is mainly used for Windows Excel automation compatibility.

---

## Run｜运行方式

```bash
python "YaFo purchase assistant_v3.5.39_Integration of the old Excel Marcos.py"
```

或者将主程序文件名改为更简洁的：

```bash
python YaFo_Purchase_Assistant.py
```

---

## Suggested Repository Structure｜建议仓库结构

```txt
YaFo-Compatible-Purchase-Assistant/
├─ YaFo_Purchase_Assistant.py
├─ README.md
├─ CHANGELOG.md
├─ requirements.txt
└─ .gitignore
```

Recommended `.gitignore`:

```txt
__pycache__/
*.pyc
*.pyo
*.log
*.xlsx
*.xls
*.xlsm
*.csv
*.xml
config_v3.txt
.po_xml_generator_cache/
```

---

## Data Security｜数据安全说明

This repository should not include real business data.

请勿上传以下文件到 GitHub：

- 真实销售订单
- 客户资料
- 供应商报价
- NC 导出数据
- 成本表
- 发票明细
- 生成后的 XML 文件
- 含真实业务信息的 Excel / CSV

If the tool contains business-specific rules or embedded templates, keeping the repository **Private** is strongly recommended.

---

## Version｜版本说明

Current version: **v3.5.39**

Recent updates:

- `v3.5.39`：新增古法采购订单 `.xlsm` 模板下载入口。
- `v3.5.38`：现代化 UI 改版，优化卡片布局、按钮 hover、日志区显示。
- `v3.5.37`：修复功能选择显示不全问题，改为多行网格布局。
- `v3.5.36`：集成销量预测与备货建议子模块。
- `v3.5.35`：集成存货成本回填子模块。
- `v3.5.34`：集成新品上新定价子模块。
- `v3.5.31 - v3.5.33`：功能选择、界面显示和供应商直发表格处理优化。
- `v3.5.28 - v3.5.30`：集成 NC 存货 CSV 导入、周结预定单处理和供应商拆分逻辑。
- `v3.1 - v3.5`：合并工厂直发与分仓订单生成逻辑，逐步修复 XML 兼容性和供应商格式识别问题。
- `v2.6`：工厂直发订单导入工具阶段。
- `v1.4`：分仓订单导入工具阶段。

---

## Project Philosophy｜项目理念

This tool is built for practical procurement execution.

It does not aim to replace ERP / NC systems. Instead, it acts as a bridge between messy real-world supplier files and strict system-import formats.

它的目标不是替代 NC 或 ERP，而是把“供应商文件 → 清洗匹配 → 采购订单 → XML导入 → 汇总核对 → 成本/预测辅助”这一段最容易耗费人工的流程标准化。

核心目标：

- Less manual copy-paste
- Fewer formatting errors
- Faster order processing
- Better traceability
- More structured procurement decisions

---

## License｜许可说明

Internal use only unless otherwise specified.

如需公开发布，建议先清理业务敏感字段、真实供应商名称、公司内部字段、模板数据和历史业务规则。
