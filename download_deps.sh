#!/usr/bin/env bash
# ============================================================
#  在「能联网、且与目标麒麟同 CPU 架构 / 同操作系统」的机器上运行
#  作用：把 PyQt5 / openpyxl / xlrd 的离线安装包下载到 ./wheels
#  之后用 U 盘把整个 wheels 目录拷到离线目标机即可。
# ============================================================
set -e

mkdir -p wheels

echo "==> 下载离线依赖包到 ./wheels"
# PyQt5 的 wheel 自带 Qt 运行库，但【跨架构不通用】，
# 必须在与目标机相同架构的机器上下载（x86_64 下 x86_64，arm64 下 arm64）。
pip download pyqt5 openpyxl xlrd -d wheels

echo "==> 完成"
echo "请把本目录下的 wheels/ 文件夹用 U 盘拷到离线目标机，"
echo "然后在目标机上执行： bash offline_install.sh"
