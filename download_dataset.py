"""
Steam 游戏数据集下载脚本
数据集: Steam Video Games 2024
来源: Kaggle (praffulsingh009/steam-video-games-2024)
数据量: ~96,500 款游戏，包含价格、评分、类型、拥有者数等 40 个字段
"""

import os
import sys

def download_with_kagglehub():
    """使用 kagglehub 下载数据集"""
    try:
        import kagglehub
    except ImportError:
        print("正在安装 kagglehub...")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "kagglehub"])
        import kagglehub

    print("正在下载 Steam Video Games 2024 数据集...")
    path = kagglehub.dataset_download("praffulsingh009/steam-video-games-2024")
    print(f"数据集下载到: {path}")

    # 复制 CSV 到当前目录
    import shutil
    csv_files = []
    for root, dirs, files in os.walk(path):
        for f in files:
            if f.endswith('.csv'):
                src = os.path.join(root, f)
                dst = os.path.join(os.getcwd(), f)
                shutil.copy2(src, dst)
                csv_files.append(dst)
                print(f"已复制: {f} -> {dst}")

    return csv_files


def download_from_kaggle_api():
    """备用方案：使用 kaggle API"""
    try:
        from kaggle.api.kaggle_api_extended import KaggleApi
        api = KaggleApi()
        api.authenticate()
        print("正在通过 Kaggle API 下载...")
        api.dataset_download_files(
            "praffulsingh009/steam-video-games-2024",
            path=os.getcwd(),
            unzip=True
        )
        print("下载完成！")
    except Exception as e:
        print(f"Kaggle API 下载失败: {e}")
        print("请尝试方法一（kagglehub），或手动下载。")


if __name__ == "__main__":
    print("=" * 60)
    print("  Steam 游戏数据集下载工具")
    print("  数据集: Steam Video Games 2024")
    print("=" * 60)

    files = download_with_kagglehub()

    if files:
        print(f"\n✅ 下载完成！共 {len(files)} 个文件。")
        for f in files:
            size_mb = os.path.getsize(f) / (1024 * 1024)
            print(f"   📄 {os.path.basename(f)} ({size_mb:.1f} MB)")
        print("\n接下来运行 analyze.py 开始数据分析！")
    else:
        print("\n⚠️ 自动下载失败，请手动下载：")
        print("   1. 访问 https://www.kaggle.com/datasets/praffulsingh009/steam-video-games-2024")
        print("   2. 点击 Download 按钮下载 CSV 文件")
        print("   3. 将 CSV 文件放到当前目录")
