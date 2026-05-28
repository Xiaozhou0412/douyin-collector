# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import shutil
import stat
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font


APP_NAME = "抖音采集工具"
INPUT_HEADER = "抖音链接，支持带口令的文本"


def project_root() -> Path:
    return Path(__file__).resolve().parent


def remove_path(path: Path) -> None:
    if path.is_dir():
        shutil.rmtree(path)
    elif path.exists():
        path.unlink()


def copy_if_exists(source: Path, destination: Path) -> None:
    if source.exists():
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)


def create_input_template(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "待采集视频链接"
    sheet["A1"] = INPUT_HEADER
    sheet["A1"].font = Font(bold=True)
    sheet["A1"].alignment = Alignment(vertical="center")
    sheet.column_dimensions["A"].width = 90
    workbook.save(path)


def write_text(path: Path, text: str, executable: bool = False) -> None:
    path.write_text(text, encoding="utf-8")
    if executable:
        mode = path.stat().st_mode
        path.chmod(mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)


def executable_name(platform_name: str) -> str:
    return f"{APP_NAME}.exe" if platform_name == "windows" else APP_NAME


def instructions(platform_name: str) -> str:
    if platform_name == "macos":
        launch_note = (
            "macOS 使用方法\n\n"
            "1. 解压后，右键点击“启动.command”，选择“打开”。\n"
            "2. 如果系统提示无法验证开发者，请在“系统设置 > 隐私与安全性”里允许打开，或再次右键打开。\n"
            "3. 第一次运行会创建“待采集视频链接.xlsx”，把抖音链接放到 A 列后再运行一次。\n"
        )
        browser_note = (
            "浏览器说明\n\n"
            "工具会自动查找 Chrome、Edge、Brave、Arc、Opera、Vivaldi、Chromium、QQ/360/搜狗浏览器、Firefox。\n"
            "如果没有找到，会使用随包 Chromium。Safari 本身不能被 Playwright 直接控制，所以不用单独安装 Safari 插件。\n"
        )
    else:
        launch_note = (
            "Windows 使用方法\n\n"
            "1. 双击“抖音采集工具.exe”运行。\n"
            "2. 第一次运行会创建“待采集视频链接.xlsx”，把抖音链接放到 A 列后再运行一次。\n"
        )
        browser_note = (
            "浏览器说明\n\n"
            "工具会自动查找 Chrome、Edge、Brave、Opera、Vivaldi、Chromium、360、QQ、搜狗、猎豹、遨游、Firefox。\n"
            "如果没有找到，会使用随包 Chromium。\n"
        )

    return (
        f"{APP_NAME} 使用说明\n\n"
        f"{launch_note}\n"
        "通用说明\n\n"
        "1. 运行后会生成“采集结果.xlsx”和 screenshots 截图文件夹。\n"
        "2. 不要只发送单个可执行文件。请发送整个压缩包，里面的 _internal、models、采集配置.txt 都要保留。\n"
        "3. “负责人/对接人”列默认留空。\n"
        "4. 目前只处理抖音链接，不处理小红书链接。\n\n"
        f"{browser_note}\n"
        "手动指定浏览器\n\n"
        "打开“采集配置.txt”，填写：\n"
        "浏览器路径=浏览器完整路径\n\n"
        "Windows 示例：\n"
        "浏览器路径=C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe\n\n"
        "macOS 示例：\n"
        "浏览器路径=/Applications/Google Chrome.app\n\n"
        "采集速度和截图尺寸\n\n"
        "“一次性采集数量”控制同时采集多少条。数量越大越快，但越容易触发验证码或加载失败。\n"
        "“失败重试一次性采集数量”建议保持 1。\n"
        "截图裁剪和 Excel 图片大小都在“采集配置.txt”里调，改完保存后重新运行。\n"
    )


def write_macos_launcher(package_dir: Path) -> None:
    launcher = package_dir / "启动.command"
    text = """#!/bin/zsh
cd "$(dirname "$0")" || exit 1
chmod +x "./抖音采集工具" 2>/dev/null
"./抖音采集工具"
echo
echo "运行结束，可以关闭这个窗口。"
read "?按回车关闭窗口。"
"""
    write_text(launcher, text, executable=True)


def prepare_package(platform_name: str) -> Path:
    root = project_root()
    package_dir = root / "dist" / APP_NAME
    if not package_dir.exists():
        raise SystemExit(f"未找到 PyInstaller 输出目录：{package_dir}")

    remove_path(package_dir / "采集结果.xlsx")
    remove_path(package_dir / "screenshots")
    remove_path(package_dir / "待采集视频链接.xlsx")

    copy_if_exists(root / "采集配置.txt", package_dir / "采集配置.txt")
    copy_if_exists(
        root / "models" / "face_detection_yunet_2023mar.onnx",
        package_dir / "models" / "face_detection_yunet_2023mar.onnx",
    )
    create_input_template(package_dir / "待采集视频链接.xlsx")
    write_text(package_dir / "使用说明.txt", instructions(platform_name))

    exe_path = package_dir / executable_name(platform_name)
    if exe_path.exists():
        exe_path.chmod(exe_path.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)

    if platform_name == "macos":
        write_macos_launcher(package_dir)

    return package_dir


def make_zip(package_dir: Path, suffix: str) -> Path:
    root = project_root()
    release_dir = root / "release"
    release_dir.mkdir(exist_ok=True)
    archive_base = release_dir / f"{APP_NAME}_{suffix}"
    zip_path = archive_base.with_suffix(".zip")
    remove_path(zip_path)
    shutil.make_archive(str(archive_base), "zip", root_dir=package_dir.parent, base_dir=package_dir.name)
    return zip_path


def parse_args() -> argparse.Namespace:
    default_platform = "macos" if sys.platform == "darwin" else "windows"
    parser = argparse.ArgumentParser(description="Prepare a sendable release zip.")
    parser.add_argument("--platform", choices=["windows", "macos"], default=default_platform)
    parser.add_argument("--suffix", default=None)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    suffix = args.suffix or args.platform
    package_dir = prepare_package(args.platform)
    zip_path = make_zip(package_dir, suffix)
    print(f"Release package: {zip_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
