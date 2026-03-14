# 项目简介 | Project Description

## 中文

**Excel 链接提取下载工具** 是一个用于从 Excel 表格中批量提取链接并使用 aria2 多线程下载的实用工具。

### 核心功能
- 三步式工作流程：配置命名 → 提取链接 → 批量下载
- 支持行/列两种命名方式，灵活配置文件名前缀
- 自动提取超链接和文本中的 URL
- 使用 aria2 多线程下载，支持断点续传
- 交互式操作界面，简单易用

### 适用场景
- 从问卷星等平台导出的 Excel 表格中批量下载图片附件
- 批量下载包含链接的任何 Excel 表格中的文件
- 需要按特定规则（如姓名、编号）重命名下载文件的场景

---

## English

**Excel Link Extract Downloader** is a practical tool for batch extracting links from Excel spreadsheets and downloading them using aria2 multi-threaded downloader.

### Key Features
- Three-step workflow: Configure Naming → Extract Links → Batch Download
- Supports row/column-based naming methods for flexible filename prefix configuration
- Automatically extracts URLs from hyperlinks and cell text
- Uses aria2 for multi-threaded downloading with resume support
- Interactive interface, easy to use

### Use Cases
- Batch download image attachments from Excel files exported from survey platforms (e.g., Sojump/Wenjuanxing)
- Batch download files from any Excel spreadsheet containing links
- Scenarios requiring downloaded files to be renamed according to specific rules (e.g., by name, ID)

---

## 技术栈 | Tech Stack

- Python 3.9+
- openpyxl (Excel 处理)
- pandas (数据处理)
- aria2 (多线程下载)

## 许可证 | License

MIT License
