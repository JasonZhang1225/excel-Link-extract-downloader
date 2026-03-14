# Excel Link Extract Downloader | Excel 链接提取下载工具

<p align="center">
  <a href="#english">English</a> | <a href="#chinese">中文</a>
</p>

---

<a name="english"></a>
## 🇺🇸 English

A practical tool for batch extracting links from Excel spreadsheets and downloading them using aria2 multi-threaded downloader.

### Features

- **Three-step workflow**: Configure Naming → Extract Links → Batch Download
- **Flexible naming**: Supports row/column-based naming methods for filename prefix configuration
- **URL extraction**: Automatically extracts URLs from hyperlinks and cell text
- **Multi-threaded download**: Uses aria2 with resume support
- **Interactive interface**: Easy to use with interactive prompts

### Prerequisites

- Python 3.9+ Recommend
- aria2 downloader

#### Install aria2

- **macOS (with Homebrew installed)**: `brew install aria2`
- **Windows**: Download from [aria2 releases](https://github.com/aria2/aria2/releases) and add to PATH
- **Linux**: `sudo apt install aria2` or `sudo yum install aria2`

### Quick Start

#### 1. Create Virtual Environment and Install Dependencies

```bash
# Create virtual environment
python3 -m venv .excelvenv

# Activate (macOS/Linux)
source .excelvenv/bin/activate

# Activate (Windows PowerShell)
.excelvenv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt
```

#### 2. Three-Step Workflow

**Step 1: Configure Naming Method**

Run `name_method_init.py` to select Excel file and configure naming:

```bash
python name_method_init.py
```

Features:
- Auto-detect `.xlsx` files in current directory
- Two naming methods:
  - **Row-based**: Each row corresponds to a person. Select a column (e.g., column with names), and the cell content in that column will be used as the filename prefix for all files in that row
  - **Column-based**: Each column corresponds to a person. Select a row (e.g., row with names), and the cell content in that row will be used as the filename prefix for all files in that column
- Generates `naming_config.json` configuration file

How it works:
1. Choose naming method: Row-based or Column-based
2. Select reference column/row: Specify which column (for row-based) or row (for column-based) contains the identifier (e.g., name, ID)
3. The tool uses the cell content from the selected column/row as the filename prefix

Example:
- Select row-based naming with column 3 (Name column) as reference → Files in row 2 will be named: `John-1.jpg`, `John-2.jpg`
- If cell C2 contains "John", all downloaded files from row 2 will start with "John-"

**Step 2: Extract Download Links**

Run `extract_links.py` to extract all links from Excel:

```bash
python extract_links.py
```

Features:
- Auto-reads `naming_config.json` configuration
- Traverses all cells to extract URLs (hyperlinks and text URLs)
- Generates filenames with prefixes based on naming config
- Outputs `aria2_links.txt` (aria2-compatible download list)

Optional parameters:
```bash
python extract_links.py "your_file.xlsx"    # Specify input file
python extract_links.py --check              # Check link accessibility
python extract_links.py --workers 20         # Set concurrent check threads
```

**Step 3: Start Downloading**

Run `finalstep_downloader.py` to download files using aria2:

```bash
python finalstep_downloader.py
```

Features:
- Auto-detects aria2c environment
- Reads `aria2_links.txt` file
- Interactive download directory selection
- Batch download with resume support

Download features:
- Multi-threaded concurrent download (default 5 threads)
- Auto-renames files according to naming configuration
- Supports resume, can be interrupted and continued anytime

### File Structure

| File | Description |
|------|-------------|
| `name_method_init.py` | Step 1: Configure naming method |
| `extract_links.py` | Step 2: Extract download links |
| `finalstep_downloader.py` | Step 3: Execute download |
| `naming_config.json` | Naming config (auto-generated) |
| `aria2_links.txt` | Download list (auto-generated) |
| `example.xlsx` | Example Excel file |

### Use Cases

- Batch download image attachments from Excel files exported from survey platforms
- Batch download files from any Excel spreadsheet containing links
- Scenarios requiring downloaded files to be renamed according to specific rules

### Notes

1. Must run Step 1 first to generate `naming_config.json` before Step 2
2. Step 2 generates `aria2_links.txt`, Step 3 depends on this file
3. Ensure aria2c is installed and added to system PATH
4. Downloaded files will be auto-renamed according to naming configuration

---

<a name="chinese"></a>
## 🇨🇳 中文

一个用于从 Excel 表格中批量提取链接并使用 aria2 多线程下载的实用工具。

### 功能特点

- **三步式工作流程**：配置命名 → 提取链接 → 批量下载
- **灵活命名**：支持行/列两种命名方式，灵活配置文件名前缀
- **URL 提取**：自动提取超链接和文本中的 URL
- **多线程下载**：使用 aria2，支持断点续传
- **交互式界面**：交互式提示，简单易用

### 前期准备

- 推荐Python 3.9+
- aria2 下载工具

#### 安装 aria2

- **macOS(已安装 Homebrew)**: `brew install aria2`
- **Windows**: 从 [aria2 releases](https://github.com/aria2/aria2/releases) 下载并添加到 PATH
- **Linux**: `sudo apt install aria2` 或 `sudo yum install aria2`

### 快速开始

#### 1. 创建虚拟环境并安装依赖

```bash
# 创建虚拟环境
python3 -m venv .excelvenv

# 激活虚拟环境（macOS/Linux）
source .excelvenv/bin/activate

# 激活虚拟环境（Windows PowerShell）
.excelvenv\Scripts\Activate.ps1

# 安装依赖
pip install -r requirements.txt
```

#### 2. 三步使用流程

**第一步：配置命名方法**

运行 `name_method_init.py` 选择 Excel 文件并配置命名方式：

```bash
python name_method_init.py
```

功能说明：
- 自动检测当前目录下的 `.xlsx` 文件
- 支持两种命名方式：
  - **行命名**：每行对应一个人。选择某一列（如姓名列），该列的单元格内容将作为对应行所有下载文件的文件名前缀
  - **列命名**：每列对应一个人。选择某一行（如姓名行），该行的单元格内容将作为对应列所有下载文件的文件名前缀
- 生成 `naming_config.json` 配置文件

工作原理：
1. 选择命名方式：行命名或列命名
2. 指定参考列/行：选择包含标识信息（如姓名、编号）的列（行命名时）或行（列命名时）
3. 工具将使用所选列/行中的单元格内容作为文件名前缀

示例：
- 选择行命名，以第3列（姓名列）作为参考列 → 第2行的文件将被命名为：`张三-1.jpg`, `张三-2.jpg`
- 如果 C2 单元格内容为"张三"，则第2行所有下载的文件名都会以"张三-"开头

**第二步：提取下载链接**

运行 `extract_links.py` 从 Excel 中提取所有链接：

```bash
python extract_links.py
```

功能说明：
- 自动读取 `naming_config.json` 配置
- 遍历 Excel 中的所有单元格，提取 URL（包括超链接和文本中的 URL）
- 根据命名配置生成带前缀的文件名
- 输出 `aria2_links.txt` 文件（aria2 可用的下载列表格式）

可选参数：
```bash
python extract_links.py "你的文件.xlsx"    # 指定输入文件
python extract_links.py --check              # 检查链接可达性
python extract_links.py --workers 20         # 设置并发检查线程数
```

**第三步：开始下载**

运行 `finalstep_downloader.py` 使用 aria2 下载文件：

```bash
python finalstep_downloader.py
```

功能说明：
- 自动检测 aria2c 环境
- 读取 `aria2_links.txt` 文件
- 交互式选择下载目录
- 使用 aria2 批量下载，支持断点续传

下载特性：
- 多线程并发下载（默认 5 线程）
- 自动按命名配置重命名文件
- 支持断点续传，可随时中断继续

### 文件说明

| 文件 | 说明 |
|------|------|
| `name_method_init.py` | 第一步：配置命名方法 |
| `extract_links.py` | 第二步：提取下载链接 |
| `finalstep_downloader.py` | 第三步：执行下载 |
| `naming_config.json` | 命名配置（自动生成） |
| `aria2_links.txt` | 下载列表（自动生成） |
| `example.xlsx` | 示例 Excel 文件 |

### 使用场景

- 从问卷星等平台导出的 Excel 表格中批量下载图片附件
- 批量下载包含链接的任何 Excel 表格中的文件
- 需要按特定规则（如姓名、编号）重命名下载文件的场景

### 注意事项

1. 必须先运行第一步生成 `naming_config.json`，才能执行第二步
2. 第二步会生成 `aria2_links.txt`，第三步依赖此文件
3. 确保 aria2c 已安装并添加到系统 PATH
4. 下载的文件会按命名配置自动重命名

---

## License | 许可证

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件。
