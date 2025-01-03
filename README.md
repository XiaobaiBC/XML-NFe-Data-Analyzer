# NFe 数据分析器 (NFe Data Analyzer)

一个用于加载单个或多个巴西电子发票(NFe)XML文件的程序，提供直观的数据分析和可以到处XLSX格式进行定制化数据分析

## 功能特点

### 📊 数据分析
- 支持批量导入 XML 格式的 NFe 文件
- 自动计算关键统计指标
- 提供详细的税收分析（ICMS、PIS、COFINS）
- 商品和客户数据统计

### 📈 数据展示
- 多页面设计，信息分类清晰
- 数据明细表格展示
- 发票分析表格
- 统计分析面板

### 💾 数据处理
- 支持导出数据到 Excel
- 智能数据过滤和搜索
- 平滑的用户界面体验

## 安装说明

### 系统要求
- Python 3.7+
- Windows/macOS/Linux

### 依赖包
bash
pip install -r requirements.txt

### 主要依赖
- tkinter
- ttkbootstrap
- openpyxl
- xml.etree.ElementTree

## 使用方法

1. 启动程序
bash
python nfe_analyzer.py


2. 选择文件
- 点击"选择NFe文件"按钮
- 选择一个或多个 XML 格式的 NFe 文件
- ![image](https://github.com/user-attachments/assets/33e54baf-6515-450d-8852-defa80972a9b)


3. 查看分析
- 在"数据明细"标签页查看详细数据
- ![image](https://github.com/user-attachments/assets/ddf1e331-f982-4ad5-9333-fadb0f94ed4c)

- 在"发票分析"标签页查看发票汇总
- ![image](https://github.com/user-attachments/assets/ac587e48-7869-4d78-b16e-0cd4c6116cb4)

- 在"统计分析"标签页查看统计指标
![image](https://github.com/user-attachments/assets/676d4e4b-0d52-442a-8c8c-44a21ea12fdb)


以下是优化后的介绍：

---

# NFe XML 分析工具：一站式发票数据分析与导出解决方案

专为处理和分析巴西电子发票（NFe）XML文件而设计，本程序提供直观的数据分析能力和高度定制化的报告导出功能。无论是从 **Bling** 等平台导出的发票，还是其他来源的销售数据，该工具都能轻松助力数据洞察与业务决策。

---

## 核心功能

### 📊 **智能数据分析**
- 批量导入 NFe XML 文件，支持多文件同时分析
- 自动生成关键统计指标
- 深入解析税收数据（如 **ICMS**、**PIS**、**COFINS** 等）
- 商品和客户数据统计，助您全面掌握业务状况

### 📈 **清晰直观的数据展示**
- 多标签页面设计，分类信息一目了然
- 数据明细表格，展示每张发票的详细内容
- 发票汇总和统计分析面板，快速掌握整体情况

### 💾 **高效的数据处理与导出**
- 一键导出分析结果到 Excel（XLSX 格式）
- 支持智能数据过滤与搜索，方便定位关键数据
- 友好的界面设计，操作流畅，无需专业技能

---

## 安装与运行

### **系统要求**
- **Python** 版本：3.7+
- 操作系统：Windows、macOS 或 Linux

### **安装依赖**
```bash
pip install -r requirements.txt
```

### **依赖包清单**
- **tkinter**：用户界面
- **ttkbootstrap**：现代化样式
- **openpyxl**：Excel 数据处理
- **xml.etree.ElementTree**：XML 文件解析

### **快速启动**
1. 启动程序：
   ```bash
   python nfe_analyzer.py
   ```
2. 导入文件：
   - 点击“选择 NFe 文件”按钮
   - 批量选择 XML 格式的 NFe 文件
   - - ![image](https://github.com/user-attachments/assets/33e54baf-6515-450d-8852-defa80972a9b)
3. 查看分析结果：
   - 在“数据明细”页查看发票详情
   - ![image](https://github.com/user-attachments/assets/ddf1e331-f982-4ad5-9333-fadb0f94ed4c)
   - 在“发票分析”页了解整体情况
   -![image](https://github.com/user-attachments/assets/ac587e48-7869-4d78-b16e-0cd4c6116cb4)
   - 在“统计分析”页获取数据洞察
   - ![image](https://github.com/user-attachments/assets/676d4e4b-0d52-442a-8c8c-44a21ea12fdb)



4. 导出数据：
   - 点击“导出 Excel”按钮，生成自定义分析报告

---

## 亮点功能

### 🎯 **深度数据洞察**
- 发票总量统计
- 商品销量与客户数据分析
- 各类税收统计
- 销售额价格区间分析

### 🛠 **技术优势**
- 现代化用户界面，简洁高效
- 快速响应的大数据处理机制
- 平滑滚动与防抖动设计，提升操作体验

---

通过该工具，您可以轻松提取、分析和导出发票数据，为业务决策提供有力支持。不论是企业主、会计师还是财务分析师，这款工具都能助您高效完成工作。立即下载并体验！
