#!/usr/bin/env bash
# ============================================================
#  在「不能联网」的麒麟目标机上运行（需先把 wheels/ 拷到同目录）
#  作用：不连网，用本地 wheels 安装 PyQt5 / openpyxl / xlrd
# ============================================================
set -e

if [ ! -d wheels ]; then
  echo "错误：未找到 wheels/ 目录，请先用 U 盘把下载好的 wheels 拷到本脚本同目录。"
  exit 1
fi

echo "==> 离线安装依赖（不从网络下载）"
pip install --no-index --find-links=./wheels pyqt5 openpyxl xlrd

echo "==> 完成，可运行程序："
echo "    python3 wps_table_mapper.py"
