#!/bin/bash
# ========================================
# WPS表格字段映射工具 —— 麒麟 Linux 一键打包脚本
# 使用方法：在麒麟 Linux 终端中执行 ./build.sh
# ========================================
set -e

echo "===== 检查 Python3 ====="
python3 --version

echo "===== 安装系统依赖（PyQt5 走 apt，避免 ARM64 无预编译 wheel）====="
sudo apt update
sudo apt install -y python3-pyqt5 python3-openpyxl python3-xlrd

echo "===== 安装打包工具（纯 Python 包，pip 装没问题）====="
pip3 install --user pyinstaller

# 确保 --user 安装的 pyinstaller 可被找到
export PATH="$HOME/.local/bin:$PATH"

echo "===== 验证依赖 ====="
python3 -c "from PyQt5.QtWidgets import QApplication; import openpyxl; import xlrd; print('依赖OK')"
if [ $? -ne 0 ]; then
    echo "✗ 依赖验证失败，请检查上方错误信息"
    exit 1
fi

echo "===== 验证 PyInstaller ====="
python3 -m PyInstaller --version
if [ $? -ne 0 ]; then
    echo "✗ PyInstaller 未安装成功，请手动运行: pip3 install --user pyinstaller"
    exit 1
fi

echo "===== 开始打包 ====="
python3 -m PyInstaller wps_table_mapper.spec --clean --noconfirm

echo "===== 打包完成 ====="
echo "输出目录: dist/WPS表格字段映射工具/"
echo "可执行文件: dist/WPS表格字段映射工具/WPS表格字段映射工具"
echo ""
echo "如需在其他麒麟机器上运行，将整个 dist/WPS表格字段映射工具/ 文件夹拷贝过去即可。"
