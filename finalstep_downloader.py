#!/usr/bin/env python3
"""
交互式 aria2 下载器
自动检测 aria2c 环境，读取 extract_links.py 生成的 aria2_links.txt 文件并开始下载
"""

import os
import subprocess
import platform


def check_aria2c():
    """检测 aria2c 是否安装并添加到环境变量"""
    print("=== 第一步：检测 aria2c 环境 ===")
    
    try:
        result = subprocess.run(
            ["aria2c", "--version"],
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0:
            version_line = result.stdout.split('\n')[0]
            print(f"✓ 检测到 aria2c: {version_line}")
            return True
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"检测 aria2c 时出错: {e}")
    
    print("\n✗ 未检测到 aria2c 或 aria2c 未添加到环境变量")
    
    if platform.system() == "Windows":
        print("\n需要安装 aria2c 并添加到环境变量:")
        print("1. 下载 aria2: https://github.com/aria2/aria2/releases")
        print('2. 解压后将 aria2c.exe 所在目录添加到系统环境变量 PATH')
        print('   - 右键"此电脑" -> "属性" -> "高级系统设置" -> "环境变量"')
        print('   - 在"系统变量"中找到"Path"，点击"编辑" -> "新建"，添加 aria2c.exe 所在目录路径')
    elif platform.system() == "Darwin":
        print("\n使用 Homebrew 安装 aria2c:")
        print("  brew install aria2")
        print("\n或者手动下载:")
        print("  下载: https://aria2.github.io/")
        print("  安装后确保 aria2c 在 PATH 中")
    else:
        print("\n使用包管理器安装:")
        print("  Ubuntu/Debian: sudo apt install aria2")
        print("  CentOS/RHEL: sudo yum install aria2")
        print("  Arch Linux: sudo pacman -S aria2")
    
    install = input("\n是否需要帮助配置 aria2c? (y/n): ").strip().lower()
    if install != 'y':
        print("请手动安装 aria2c 后重新运行本脚本。")
        return False
    
    return install_aria2c()


def install_aria2c():
    """帮助用户安装 aria2c"""
    system = platform.system()
    
    if system == "Darwin":
        print("\n尝试使用 Homebrew 安装 aria2c...")
        try:
            result = subprocess.run(["brew", "install", "aria2"], capture_output=True, text=True)
            if result.returncode == 0:
                print("✓ aria2c 安装成功！")
                return True
            else:
                print("✗ Homebrew 安装失败，请手动安装")
        except FileNotFoundError:
            print("✗ 未找到 Homebrew，请先安装 Homebrew: https://brew.sh/")
    elif system == "Windows":
        print("\n请手动下载并安装 aria2c:")
        print("1. 下载: https://github.com/aria2/aria2/releases")
        print('2. 解压后将 aria2c.exe 所在目录添加到系统环境变量 PATH')
        print('   - 右键"此电脑" -> "属性" -> "高级系统设置" -> "环境变量"')
        print('   - 在"系统变量"中找到"Path"，点击"编辑" -> "新建"，添加 aria2c.exe 所在目录路径')
    else:
        print("\n请使用包管理器安装:")
        print("  sudo apt install aria2  # Ubuntu/Debian")
        print("  sudo yum install aria2  # CentOS/RHEL")
        print("  sudo pacman -S aria2  # Arch Linux")
    
    return False


def check_aria2_file():
    """检查 aria2_links.txt 文件是否存在"""
    print("\n=== 第二步：检查 aria2_links.txt 文件 ===")
    
    aria2_file = "aria2_links.txt"
    
    if os.path.exists(aria2_file):
        print(f"✓ 找到 aria2_links.txt 文件")
        
        with open(aria2_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            url_count = len([l for l in lines if l.strip() and not l.startswith('  out=')])
            print(f"  共 {url_count} 个下载链接")
        
        return aria2_file
    else:
        print(f"✗ 未找到 aria2_links.txt 文件")
        print("\n请先运行 extract_links.py 生成下载链接文件:")
        if platform.system() == "Windows":
            print('  python extract_links.py --use-naming-config')
        else:
            print('  python3 extract_links.py --use-naming-config')
        print("\n生成 aria2_links.txt 后再运行本脚本。")
        return None


def get_download_directory():
    """询问用户文件保存位置"""
    print("\n=== 第三步：选择文件保存位置 ===")
    print("请输入下载文件的保存目录")
    print("提示: 可以直接拖入需要保存到的文件夹，或直接回车使用默认目录")
    
    # 默认下载目录：项目文件夹/DownloadedFiles
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_dir = os.path.join(script_dir, "DownloadedFiles")
    
    download_dir = input(f"保存目录 (默认: {default_dir}): ").strip()
    
    if not download_dir:
        download_dir = default_dir
        print(f"使用默认目录: {download_dir}")
    else:
        download_dir = os.path.expanduser(download_dir)
        download_dir = os.path.abspath(download_dir)
        
        if not os.path.exists(download_dir):
            print(f"目录不存在，正在创建: {download_dir}")
            try:
                os.makedirs(download_dir, exist_ok=True)
                print(f"✓ 目录创建成功: {download_dir}")
            except Exception as e:
                print(f"✗ 目录创建失败: {e}")
                return None
        else:
            print(f"✓ 使用目录: {download_dir}")
    
    return download_dir


def start_download(aria2_file, download_dir):
    """开始下载"""
    print("\n=== 开始下载 ===")
    
    abs_aria2_file = os.path.abspath(aria2_file)
    abs_download_dir = os.path.abspath(download_dir)
    
    print(f"aria2 列表文件: {abs_aria2_file}")
    print(f"下载目录: {abs_download_dir}")
    
    confirm = input("\n确认开始下载? (y/n): ").strip().lower()
    if confirm != 'y':
        print("下载已取消")
        return False
    
    print("\n正在启动 aria2c...")
    
    cmd = [
        "aria2c",
        "-i", abs_aria2_file,
        "-d", abs_download_dir,
        "--allow-overwrite=true",
        "--auto-file-renaming=false",
        "-j", "5",
        "-x", "5",
        "-s", "5"
    ]
    
    print(f"执行命令: {' '.join(cmd)}")
    print("\n下载过程中可以按 Ctrl+C 暂停，下载进度会自动保存")
    print("重新运行本脚本可以继续下载未完成的文件\n")
    
    try:
        result = subprocess.run(cmd)
        if result.returncode == 0:
            print("\n✓ 所有文件下载完成！")
            return True
        else:
            print("\n✗ 下载过程中出现错误，请检查输出信息")
            return False
    except KeyboardInterrupt:
        print("\n\n下载已暂停")
        print("提示: 下载进度会自动保存，重新运行本脚本可以继续下载")
        return False
    except Exception as e:
        print(f"\n✗ 启动 aria2c 失败: {e}")
        return False


def main():
    print("=" * 60)
    print("交互式 aria2 下载器")
    print("=" * 60)
    
    if not check_aria2c():
        return
    
    aria2_file = check_aria2_file()
    if not aria2_file:
        return
    
    download_dir = get_download_directory()
    if not download_dir:
        return
    
    if not start_download(aria2_file, download_dir):
        return
    
    print("\n" + "=" * 60)
    print("下载完成！")
    print("=" * 60)


if __name__ == "__main__":
    main()
