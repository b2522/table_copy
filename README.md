# WPS表格字段映射工具

单文件 PyQt5 桌面应用，用于把「WPS 源表格」的字段按自定义规则映射并写入「WPS 目标表格」。

## 功能

- **表格映射**：支持直接复制、合并复制（2~5 源列 + 连接符）、拆分复制（1 源列 → N 目标列）
- **表格合并**：多个 Excel 文件合并为一个，按指定列排序输出
- **日期处理**：自动识别日期单元格并转换为 `YYYY.MM` 格式
- **格式保留**：目标表写入时保留原格式/样式/公式，仅覆盖数据区
- **配置导入导出**：保存/恢复完整的映射配置（含文件路径、表头行、映射规则）
- **拖拽排序**：映射行支持拖拽重排序

## 支持格式

| 格式 | 读取 | 写入 |
|------|------|------|
| `.xlsx` / `.xlsm` | openpyxl | openpyxl（保留原格式） |
| `.xls` | xlrd | 转存为 `.xlsx`（不保留原格式） |

## 依赖安装

### Windows

```bash
pip install PyQt5 openpyxl xlrd
```

### 麒麟 Linux / ARM64

```bash
sudo apt update
sudo apt install -y python3-pyqt5 python3-openpyxl python3-xlrd
```

> 麒麟/ARM64 平台 PyPI 无 PyQt5 预编译 wheel，必须用 apt 安装。

## 运行

```bash
python3 wps_table_mapper.py
```

## 打包

### 麒麟 Linux 本地打包

```bash
chmod +x build.sh
./build.sh
```

产物：`dist/WPS表格字段映射工具`（单文件，可直接拷贝到其他麒麟机器运行）

### GitHub Actions 自动打包

| 平台 | Workflow 文件 | 产物 |
|------|--------------|------|
| 麒麟 Linux ARM64 | `.github/workflows/build-arm64-kylin.yml` | `WPS表格字段映射工具`（ELF 单文件） |
| Windows | `.github/workflows/build-windows.yml` | `WPS表格字段映射工具.exe` |

打 `v*` 标签推送即触发两个平台同时打包并发布 Release：

```bash
git tag v1.0.0
git push origin v1.0.0
```

也可在 GitHub 仓库 Actions 页面手动触发（workflow_dispatch）。

## 使用说明

### 表格映射

1. 点击「上传 WPS 源表格」和「上传 WPS 目标表格」加载文件
2. 根据需要调整「源表表头行」和「目标表表头行」的行号
3. 在映射配置表中逐行设置：
   - **填入方式**：直接复制 / 合并复制 / 拆分复制
   - **源列**：选择源表格的列名（合并复制可选 2~5 个）
   - **目标列**：选择目标表格的列名（拆分复制可选多个）
   - **连接符**：合并复制时各源列之间的分隔符
   - **拆分数**：拆分复制时目标列的数量
4. 点击「整表复制」执行映射并保存

### 表格合并

1. 切换到「表格合并」标签页
2. 点击「添加 WPS 表格」选择多个文件
3. 设置表头行、排序列、排序方式
4. 点击「合并并排序」预览结果
5. 点击「保存结果」导出合并后的 `.xlsx` 文件

### 配置导入导出

- **导出配置**：保存当前完整的映射配置（含文件路径、表头行号、全部映射规则）为 JSON 文件
- **导入配置**：从 JSON 文件恢复全部状态，自动加载源表/目标表并重建映射规则

## 文件结构

```
table_copy/
├── wps_table_mapper.py          # 主程序（单文件）
├── wps_table_mapper.spec        # PyInstaller 打包配置
├── build.sh                     # 麒麟 Linux 本地打包脚本
├── requirements.txt             # Python 依赖（pip 安装）
├── README.md                    # 本文件
└── .github/workflows/
    ├── build-arm64-kylin.yml    # 麒麟 Linux ARM64 打包
    └── build-windows.yml        # Windows 打包
```
