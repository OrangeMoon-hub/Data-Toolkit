# 名创数据处理工具箱

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![Jupyter](https://img.shields.io/badge/Jupyter-Notebook-orange.svg)](https://jupyter.org/)
[![Pandas](https://img.shields.io/badge/Pandas-2.x-150458.svg)](https://pandas.pydata.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

> 日常物流/人力资源数据处理工作的自动化脚本集。以 Jupyter Notebook 为载体，按步骤交互式执行，兼顾灵活性与可审计性。

---

## 📋 项目定位

本仓库记录了实际工作中沉淀的数据处理工具。两个脚本分别解决物流绩效汇总和离职数据拆分的批量 Excel 处理需求，展示了 **pandas 数据处理、多 sheet Excel 读写、正则文件筛选、分组输出** 等常见企业场景技能。

> ⚠️ 出于数据安全考虑，所有 Notebook 输出已清空，源数据路径已脱敏。运行时请按提示输入实际路径。

---

## 🧰 工具列表

### 1. `wuliu_merge.ipynb` — 物流绩效数据合并

| 项目 | 说明 |
|------|------|
| **场景** | 每月需汇总 18 个仓库的绩效核算明细表，合并为统一的集团绩效表 |
| **输入** | 各仓库提交的 `物流绩效核算明细表.xlsx`（分布在按仓分组的子目录中） |
| **输出** | `efficiency_list.xlsx`（人员绩效明细，40+ 字段）、`ware_list.xlsx`（仓库指标汇总） |
| **核心操作** | 正则匹配文件名 → 提取仓库指标 → 多 Sheet 合并 → 历史名称标准化 → 列名清洗 → 分仓复核输出 |

#### 处理流程

```
源目录 (各仓文件夹)
┌─ B1仓/ 物流绩效核算明细表.xlsx
├─ B2仓/ 物流绩效核算明细表.xlsx
├─ ...  (18个仓库)
└─ 西安仓/ 物流绩效核算明细表.xlsx
        │
        ▼ 正则筛选（含"绩效"关键字的文件，排除薪资汇总表）
        │
        ▼ 逐文件处理
        ├── 提取仓库指标（年月/是否国际/效率值等）
        ├── 读取效率数据 Sheet（配货金额）
        ├── 国际仓额外读取协助岗 Sheet
        └── 多表合并（e1 效率表 + e2 指标 + e3 金额）
        │
        ▼
  输出: efficiency_list.xlsx
        ware_list.xlsx
```

#### 技术要点

- 正则表达式 `(?=.*绩效)[\u4e00-\u9fa5]+[-仓]` 匹配中文仓库名
- `pd.read_excel(skiprows=8)` 跳过固定行数读取数据体
- `pd.read_excel(nrows=5, usecols=[...], header=None)` 精准提取文件头部元数据
- 历史仓库名称映射（潮玩仓 → TOPTOY仓）
- 列名换行符清理 `\n` → 标准列名
- 最终结果按 `统一仓库名称` 拆分为独立文件供业务复核

---

### 2. `extract.ipynb` — 离职数据按部门拆分

| 项目 | 说明 |
|------|------|
| **场景** | 集团离职明细总表按一级部门拆分为独立 Excel，分发给各业务线 BP |
| **输入** | 集团离职明细表（含 4 个 Sheet：离职/在职两个时点/BP对应表） |
| **输出** | 按部门输出：`{部门名称}-{BP姓名}.xlsx`（每个含 3 个 Sheet） |
| **核心操作** | 多 Sheet 读取 → 日期格式统一 → 分组 → openpyxl 格式化 → 单部门输出 |

#### 处理流程

```
集团离职明细.xlsx
├── Sheet「离职」        → 期间内离职人员明细
├── Sheet「202412在职」  → 2024年12月在职快照
├── Sheet「202508在职」  → 2025年8月在职快照
└── Sheet「BP对应表」    → 事业部 ↔ BP姓名映射
        │
        ▼ 按「一级部门」分组
        │
        ▼ 逐部门输出
        ├── 日期列 → yyyy/mm/dd 格式
        ├── 工号列 → 文本格式（防科学计数法）
        └── 文件名 → {部门}-{BP姓名}.xlsx
```

#### 技术要点

- `pd.to_datetime(format='%Y%m%d', errors='coerce')` 日期格式统一
- `pd.ExcelWriter` + `openpyxl` 工作表格式精细控制
- `worksheet.cell().number_format = '@'` 防工号变成科学计数法
- BP 对应表 → dict → 文件名自动拼 BP 姓名

---

## 🚀 使用方式

### 环境要求

```bash
pip install pandas numpy openpyxl
```

Python 3.8+ 均可运行，推荐使用 Jupyter Notebook / JupyterLab。

### 运行步骤

1. 用 Jupyter 打开 `.ipynb` 文件
2. 按提示在 `input()` 弹窗中输入源文件/目录路径
3. 逐 Cell 执行（Shift+Enter），每个 Cell 含独立的逻辑步骤
4. 输出文件生成在同目录或指定输出目录

### 项目结构

```
mingchuang/
├── wuliu_merge.ipynb        # 物流绩效合并
├── dwd156_extract.ipynb     # 离职数据拆分
├── requirements.txt         # 依赖清单
├── .gitignore               # 忽略规则
├── LICENSE                  # MIT 协议
└── README.md                # 本文件
```

---

## ⚠️ 注意事项

- 两个脚本均依赖本地 Excel 源文件，运行前请确保路径正确
- `wuliu_merge.ipynb` 文件名筛选依赖「绩效」关键字，请勿随意修改源文件命名规范
- `dwd156_extract.ipynb` 依赖 `一线部门` 字段存在；3 个 Sheet 全部含该字段方可拆分
- 出错时会打印失败文件路径至控制台，可根据输出排查

---

## 👤 作者

**OrangeMoon** (橙月君)

- GitHub: [@OrangeMoon-hub](https://github.com/OrangeMoon-hub)

---

## 📄 License

MIT © OrangeMoon
