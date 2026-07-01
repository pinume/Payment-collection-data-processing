# 回款数据处理工具

用于处理家电和数码补贴明细：转换原始文件、按商户编号筛选、统一明细结构、识别财务大类与品牌，并生成 Excel 整合明细和汇总表。

## 从远程仓库使用

安装 [Git](https://git-scm.com/)、[Python 3.14+](https://www.python.org/) 和 [uv](https://docs.astral.sh/uv/)，然后执行：

```powershell
git clone https://github.com/pinume/Payment-collection-data-processing.git
Set-Location Payment-collection-data-processing
uv sync --locked
uv run python main.py
```

## 运行流程

程序会依次提示：

1. 选择数据类型：家电或数码。
2. 输入要筛选的商户编号。
3. 输入原始数据文件路径。

支持 `.xlsx`、`.xls`、`.xlsm`、`.csv` 和 `.tsv`；支持普通 Windows 路径、带引号的拖入路径和 `file:///C:/...` 地址。

## 输出

结果保存到项目的 `output` 目录：

- 家电：`output\家电回款明细.xlsx`
- 数码：`output\数码回款明细.xlsx`

原始数据不会被覆盖；处理完成后，中间工作副本会自动删除。

## 更新项目

```powershell
git pull --ff-only
uv sync --locked
```
