# Steam 游戏市场数据分析

> 数据思维与人工智能 · 课程大作业  
> 基于 Steam Video Games 2024 数据集（97,428 款游戏，40 个字段）

## 项目结构

```
├── Steam游戏市场数据分析.ipynb    ← Jupyter Notebook（交互式分析）
├── analyze.py                     ← Python 脚本版（命令行运行）
├── generate_ppt.py                ← 生成演讲 PPT
├── generate_report.py             ← 生成 Word 分析报告
├── download_dataset.py            ← 从 Kaggle 下载数据集
├── 演讲稿_Steam游戏市场数据分析.md
├── 分析报告.txt
├── Steam游戏市场数据分析_演讲.pptx
│
├── 01_综合概览.png                ← 分析图表
├── 01_相关性热力图.png
├── 02_定价策略分析.png
├── 03_类型分析.png
├── 04_评价分析.png
│
└── Steam Games 2024.csv           ← 数据集（需下载，不纳入版本控制）
```

## 快速开始

### 1. 获取数据集

```bash
python download_dataset.py
```

或手动从 [Kaggle](https://www.kaggle.com/datasets/praffulsingh009/steam-video-games-2024) 下载 `Steam Games 2024.csv` 放入项目根目录。

### 2. 安装依赖

```bash
pip install pandas numpy matplotlib seaborn jupyterlab
# 可选：生成 PPT 和 Word 报告
pip install python-pptx python-docx
```

### 3. 运行分析

**方式一：Jupyter Notebook（推荐）**
```bash
jupyter lab
# 打开 Steam游戏市场数据分析.ipynb，按 Shift+Enter 逐单元格运行
```

**方式二：命令行脚本**
```bash
python analyze.py
# 生成 5 张 PNG 图表 + 分析报告.txt
```

### 4. 生成报告

```bash
python generate_report.py   # → Steam游戏市场数据分析报告.docx
python generate_ppt.py      # → Steam游戏市场数据分析_演讲.pptx
```

## 分析维度

| 维度 | 内容 |
|------|------|
| 一 | 综合评价：价格/好评率/拥有者分布、相关性矩阵 |
| 二 | 定价策略：价格区间对比、免费 vs 付费 |
| 三 | 类型趋势：Indie/动作/休闲等类型的市场份额与好评率 |
| 四 | 玩家评价：评价数 vs 好评率、Metacritic 对比、影响因素 |

## 核心发现

- 价格与质量几乎无关（r=0.046），Steam 为优质低价游戏提供公平竞争环境
- $5-$15 是定价甜蜜点，好评率最高（83.2%）
- 付费游戏好评率（82.1%）显著优于免费（76.3%），但免费用户覆盖更广
- Indie 独立游戏数量碾压（64,499 款），Casual 休闲类好评率最高
- 平台支持、DLC、成就系统与好评率正相关——持续运营带来正向反馈

## 技术栈

Python · pandas · numpy · matplotlib · seaborn · Jupyter Notebook

## 数据来源

Kaggle — [Steam Video Games 2024](https://www.kaggle.com/datasets/praffulsingh009/steam-video-games-2024)
