#!/usr/bin/env bash
# ============================================================
#  WPS表格字段映射工具 —— 麒麟Linux 打包脚本
#  用法：把整个项目文件夹（含 wps_table_mapper.py）拷到麒麟系统，
#        然后在该目录执行：  bash build_kylin.sh
#  产物：dist/WPS表格字段映射工具  （单个可执行文件，可直接双击运行）
# ============================================================
set -e

APP_NAME="WPS表格字段映射工具"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ENTRY="$SCRIPT_DIR/wps_table_mapper.py"

if [ ! -f "$ENTRY" ]; then
  echo "错误：未找到 $ENTRY，请把本脚本与 wps_table_mapper.py 放在同一目录。"
  exit 1
fi

echo "==> [1/4] 准备 Python 虚拟环境"
python3 -m venv venv || {
  echo "提示：venv 不可用，尝试安装 python3-venv ..."
  sudo apt-get update && sudo apt-get install -y python3-venv python3-pip
  python3 -m venv venv
}
# shellcheck disable=SC1091
source venv/bin/activate

echo "==> [2/4] 安装依赖（PyQt5 / openpyxl / xlrd / PyInstaller）"
pip install --upgrade pip
pip install pyqt5 openpyxl xlrd pyinstaller

echo "==> [3/4] 使用 PyInstaller 打包"
# --onefile : 打包成单个可执行文件
# --windowed: GUI 程序，不弹终端（如需看报错日志，去掉此参数重新打包）
# --collect-all PyQt5 / openpyxl : 确保 Qt 插件与 Excel 支持库完整打入
pyinstaller --noconfirm --onefile --windowed \
  --name "$APP_NAME" \
  --collect-all PyQt5 \
  --collect-all openpyxl \
  --hidden-import=xlrd \
  "$ENTRY"

echo "==> [4/4] 完成"
ls -lh "dist/$APP_NAME"
echo
echo "提示："
echo "  - 把 dist/$APP_NAME 拷到任意同架构(amd64/aarch64)的麒麟机器即可运行。"
echo "  - 若双击无反应，终端执行 ./dist/$APP_NAME 查看报错。"
echo "  - 若提示 'GLIBC_xxx not found'，说明打包机的 glibc 比目标机新，"
echo "    请直接在目标麒麟系统上重新跑本脚本即可解决。"
