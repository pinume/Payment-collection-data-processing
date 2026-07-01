### 使用说明

请先安装以下工具：

* [Git](https://git-scm.com/)
* [uv](https://docs.astral.sh/uv/)

> Python 由 uv 自动安装和管理，无需提前单独安装 Python。

然后在 PowerShell 中执行：

```powershell
git clone https://github.com/pinume/Payment-collection-data-processing.git
Set-Location Payment-collection-data-processing
uv sync --locked
uv run python main.py
```

如果需要手动安装项目要求的 Python 版本，也可以执行：

```powershell
uv python install 3.14
```

## 运行环境

当前项目运行环境如下：

* 操作系统：Windows（推荐 Windows 10 / 11）
* Python：3.14（由 uv 自动管理）
* 依赖管理工具：uv
* 主要依赖库：通过 `uv sync` 自动安装（详见项目依赖配置文件）

## 运行流程

程序启动后，会依次提示输入以下内容：

* 选择数据类型：家电或数码
* 输入要筛选的商户编号
* 输入原始数据文件路径

支持的原始文件格式：

* `.xlsx`
* `.xls`
* `.xlsm`
* `.csv`
* `.tsv`

支持的文件路径形式：

* 普通 Windows 路径
* 带引号的拖入路径
* `file:///C:/...` 地址

## 输出结果

处理结果会保存到项目的 `output` 目录。

| 数据类型 | 输出文件                 |
| ---- | -------------------- |
| 家电   | `output\家电回款明细.xlsx` |
| 数码   | `output\数码回款明细.xlsx` |

说明：

* 原始数据不会被覆盖
* 处理完成后，中间工作副本会自动删除

## 更新项目

进入项目目录后执行：

```powershell
git pull --ff-only
uv sync --locked
```
