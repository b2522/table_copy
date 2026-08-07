# WPS 表格字段映射工具 —— ARM64 麒麟打包指南

将本项目（含 `wps_table_mapper.py`）推送到 GitHub 后，利用 GitHub Actions
在 **ARM64 + 麒麟 V10 同源环境** 下用 PyInstaller 打包成单文件二进制。

## 文件清单

| 文件 | 作用 |
|------|------|
| `.github/workflows/build-arm64-kylin.yml` | 打包工作流 |
| `requirements.txt` | 纯 Python 依赖（openpyxl / xlrd / pyinstaller） |
| `wps_table_mapper.py` | 应用源码（需与本文件同仓库根目录） |

## 为什么这么设计

- **目标机**：Kirin 9006C（鲲鹏 ARM64）+ 麒麟 V10（基于 Ubuntu 20.04，glibc 2.31）。
- **Python 3.8**：用 `arm64v8/ubuntu:20.04` 容器，其自带 Python 3.8，glibc 2.31，
  与麒麟 V10 一致，避免高版本 glibc 导致的 `GLIBC_2.xx not found`。
- **PyQt5 走 apt 而非 pip**：PyPI 上 PyQt5 **没有 Linux ARM64 的 wheel**，
  pip 会回退源码编译（需 Qt qmake + 全量构建，QEMU 下极慢且易失败）。
  改用 `apt install python3-pyqt5`，装的是麒麟同源的 ARM64 编译版。

## 使用方法

1. 把 `wps_table_mapper.py`、`requirements.txt`、`.github/` 放进同一个 Git 仓库根目录。
2. 推送到 GitHub。
3. 触发构建：
   - **方式 A（推荐，出 release）**：打标签 `v1.0.0` 并推送
     ```
     git tag v1.0.0
     git push origin v1.0.0
     ```
   - **方式 B（手动）**：仓库 → Actions → 选本工作流 → Run workflow。
4. 构建产物：
   - Actions 页面的 **Artifacts** 里下载 `wps_table_mapper-arm64-kylin`；
   - 若打了 tag，还会自动在 **Releases** 里附上该二进制。

> 提示：QEMU 模拟 ARM64 较慢，单次构建约 **10–20 分钟**，请耐心等待。

## 目标机运行

将二进制拷到麒麟机器，赋予执行权限后直接运行（需桌面会话）：

```bash
chmod +x wps_table_mapper-arm64-kylin
./wps_table_mapper-arm64-kylin
```

二进制已内置 Qt 库，仅依赖系统基础图形库，麒麟桌面默认具备。
若通过 SSH 无显示环境运行，需装 `xvfb` 并用 `xvfb-run` 启动。

### 首次运行白屏 / 报 xcb 插件缺失时

补装 Qt 的平台插件依赖即可：

```bash
sudo apt install -y libxcb-xinerama0 libxcb-icccm4 libxcb-image0 \
  libxcb-keysyms1 libxcb-randr0 libxcb-render-util0 libxcb-shape0 libxcb-cursor0
```

### 中文显示

麒麟桌面通常已带文泉驿 / 思源等中文字体，程序启动时会自动探测并回退，无需额外配置。
