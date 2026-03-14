# Excel 链接提取下载工具

从 Excel 表格中批量提取链接并使用 aria2 下载的工具。

## 功能特点

- 三步式工作流程，简单高效
- 支持行/列两种命名方式，灵活配置文件名
- 自动提取超链接和文本中的 URL
- 使用 aria2 多线程下载，支持断点续传
- 交互式操作，无需记忆复杂命令

## 前期准备

- 推荐使用Python 3.9+ 环境
- aria2 下载工具

### 安装 aria2

- **macOS**: 预先安装了 homebrew 的情况下，直接执行 `brew install aria2`
- **Windows**: 从 [aria2 releases](https://github.com/aria2/aria2/releases) 下载并添加到 PATH
- **Linux**: `sudo apt install aria2` 或 `sudo yum install aria2`

## 快速开始

### 1. 创建虚拟环境并安装依赖

注意首先打开终端，进入到本项目目录

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

***

## 三步使用流程

### 第一步：配置命名方法

运行 `name_method_init.py` 选择 Excel 文件并配置命名方式：

```bash
python name_method_init.py
```

**功能说明：**

- 自动检测当前目录下的 `.xlsx` 文件
- 支持两种命名方式：
  - **行命名**：每行对应一个人，使用指定列的内容作为文件名前缀
  - **列命名**：每列对应一个人，使用指定行的内容作为文件名前缀
- 生成 `naming_config.json` 配置文件

**示例：**

- 选择行命名，姓名在第3列 → 生成的文件名格式：`张三-1.jpg`, `张三-2.jpg`

***

### 第二步：提取下载链接

运行 `extract_links.py` 从 Excel 中提取所有链接：

```bash
python extract_links.py
```

**功能说明：**

- 自动读取 `naming_config.json` 配置
- 遍历 Excel 中的所有单元格，提取 URL（包括超链接和文本中的 URL）
- 根据命名配置生成带前缀的文件名
- 输出 `aria2_links.txt` 文件（aria2 可用的下载列表格式）

**可选参数：**

```bash
python extract_links.py "你的文件.xlsx"    # 指定输入文件
python extract_links.py --check              # 检查链接可达性
python extract_links.py --workers 20         # 设置并发检查线程数
```

***

### 第三步：开始下载

运行 `finalstep_downloader.py` 使用 aria2 下载文件：

```bash
python finalstep_downloader.py
```

**功能说明：**

- 自动检测 aria2c 环境
- 读取 `aria2_links.txt` 文件
- 交互式选择下载目录
- 使用 aria2 批量下载，支持断点续传

**下载特性：**

- 多线程并发下载（默认 5 线程）
- 自动按命名配置重命名文件
- 支持断点续传，可随时中断继续

***

## 文件说明

| 文件                        | 说明          |
| ------------------------- | ----------- |
| `name_method_init.py`     | 第一步：配置命名方法  |
| `extract_links.py`        | 第二步：提取下载链接  |
| `finalstep_downloader.py` | 第三步：执行下载    |
| `naming_config.json`      | 命名配置（自动生成）  |
| `aria2_links.txt`         | 下载列表（自动生成）  |
| `example.xlsx`            | 示例 Excel 文件 |

## 完整示例流程

```bash
# 1. 进入项目目录
cd excel-Link-extract-downloader

# 2. 激活虚拟环境
source .excelvenv/bin/activate

# 3. 第一步：配置命名
python name_method_init.py
# → 选择 example.xlsx
# → 选择行命名
# → 选择第3列作为姓名列

# 4. 第二步：提取链接
python extract_links.py
# → 生成 aria2_links.txt

# 5. 第三步：下载文件
python finalstep_downloader.py
# → 选择保存目录
# → 开始下载
```

## 注意事项

1. 必须先运行第一步生成 `naming_config.json`，才能执行第二步
2. 第二步会生成 `aria2_links.txt`，第三步依赖此文件
3. 确保 aria2c 已安装并添加到系统 PATH
4. 下载的文件会按命名配置自动重命名

## License

MIT License
