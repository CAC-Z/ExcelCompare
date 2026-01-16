# Excel Compare

用于对比两个 Excel 文件差异的桌面小工具（PyQt5）。选择文件、设置索引列、配置对比列，即可导出差异报告。

## 截图

![App Screenshot](icons/screenshot.png)

## 功能特性

- 图形界面操作，适合日常对比
- 快速读取表头，避免大文件卡顿
- 列名筛选、手动映射、同名自动匹配
- 统一归一化，减少 0/0.0、空格等误差导致的误报
- 导出结果（差异汇总 + 详细对比）

## 运行环境

- Python 3.9+
- Windows（已测试）

安装依赖：

```bash
pip install -r requirements.txt
```

## 使用方式

```bash
python excel_compare.py
```

## 示例数据

示例文件位于 `examples/`：

- `examples/example_file1.xlsx`
- `examples/example_file2.xlsx`

可直接用这两个文件进行试跑，索引列选择 `ID`，对比列选择 `Name`、`Score`、`Dept`。

## 打包（PyInstaller）

```powershell
pwsh ./build.ps1 -Mode onefile -NoUPX
```

## 项目结构

- `excel_compare.py` 主程序
- `build.ps1` 打包脚本
- `ExcelCompare.spec` PyInstaller 配置
- `icons/` UI 资源
- `screenshots/` README 截图
- `examples/` 示例数据

## License

MIT，见 `LICENSE`。
