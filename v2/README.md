# Refilling Tool v2

夜间液流分解 (nighttime sap flow partitioning) 工具，PWRlab（北京林业大学）。
方法原理见 [CJPE 2023.0043](https://www.plant-ecology.com/CN/10.17521/cjpe.2023.0043)。

v2 在本地运行：数据不上传任何服务器，浏览器只连接本机 `localhost`。Mac 和 Windows 都可用。

## 与 v1 相比

| | v1 | v2 |
|---|---|---|
| 运行平台 | Windows exe | Mac / Windows / Linux，本地浏览器界面 + 命令行 |
| 输入 | xlsx，列名固定 | xlsx / csv，自动识别时间、辐射、样树列，可手动修改 |
| 夜间划分 | 首尾不完整夜晚也参与计算；每夜少算日出前最后一个点 | 完整夜晚；自动标记不完整夜晚和缺测夜晚；可忽略夜间辐射噪声 |
| 数据检查 | 无 | 缺测、重复时间、时间漂移、辐射异常值提示 |
| 结果检查 | 只能看 Excel | 逐夜查看实测曲线、各方法拟合曲线和断点；可一键排除阴雨天 |
| 晴天筛选 | 手动 | 按当日辐射总量 (MJ m⁻²) 自动筛选 |
| 输出 | 宽表 | 长表 (long format，直接用于 R/ggplot2) + v1 风格宽表 + 参数记录 |
| 依赖 | scipy==1.7.3（新版 Python 装不上） | 兼容当前 scipy / Python 3.9–3.13 |
| 可复现性 | — | `--legacy` 模式与 v1 结果逐值一致（有自动测试） |

四种方法：Re (time segment)、Exp (exponential decay)、Line (linear decay)、Et (transpiration inversion)。`fit_sum` 统一为夜间充水量 (nocturnal refilling)，`remainder` 为夜间蒸腾量 (nocturnal transpiration)。

> **Et 法已按原文修正。** 原文 1.4.2(4)：以 Re 法拐点 P2 为后半夜起点，后半夜液流最小值作为整夜蒸腾速率；蒸腾 = 速率 × 整夜时长，充水 = 夜间液流 − 蒸腾。v1 代码与此不符：使用自己的断点（10 点而非 P2），且输出的是“P2 前实测液流 + 蒸腾基线”，混合了两个组分。在示例数据上，两者逐夜充水比例的绝对差中位数为 0.32。Re、Exp、Line 与 v1 一致。`--legacy` 模式仍保留 v1 的 Et 算法，用于复现旧结果。

## 安装与启动

需要 Python ≥ 3.9。

**双击启动**（首次运行自动安装依赖，约 1–3 分钟）：
- Mac：双击 `start_mac.command`（若提示无法打开：右键 → 打开）
- Windows：双击 `start_windows.bat`

浏览器会自动打开 `http://localhost:8501`。

**命令行**：

```bash
cd v2
python3 -m pip install -e .
refilltool gui                                  # 图形界面
refilltool run data.xlsx -o results             # 四种方法全部计算
refilltool run data.csv -m refilling exp --fitting
refilltool run data.xlsx --min-day-rad 15       # 只保留当日辐射 >= 15 MJ m-2 的夜晚
refilltool run data.xlsx --legacy               # 复现 v1 结果
refilltool init-config my.json                  # 生成参数文件，修改后用 -c my.json
```

v1 的 `config.json` 可直接用 `-c config.json` 读取。

## 数据格式

一列时间、一列太阳辐射 (W m⁻²)、其余每列一棵样树的液流。列名任意，常见写法会被自动识别：

| time | solar_radiation | Tree1 | Tree2 |
|---|---|---|---|
| 2024-06-01 00:00 | 0 | 0.012 | 0.020 |
| 2024-06-01 00:10 | 0 | 0.011 | 0.019 |

时间步长自动检测（5 min、10 min、30 min 均可）。缺测记录会自动补成空值，含空值的夜晚不参与计算并在结果中注明原因。

## 参数

| 参数 | 默认 | 说明 |
|---|---|---|
| `default_point` | Re 30；其他 10 | 断点 (breakpoint)，即夜间开始后的点数。界面上按检测到的步长换算；10 min 数据的 30 点 = 5 h |
| `kneed_process` | false | 自动检测拐点 (Kneedle)。仅适用于夜间数据质量高的情况 |
| `min_strategy` | true | Et 法蒸腾速率：true = 后半夜最小值（原文）；false = 后半夜均值 |
| `threshold` | 5 W m⁻² | 辐射 ≤ 阈值为夜间 |
| `min_run_minutes` | 30 | 短于该时长的昼夜切换被忽略（夜间辐射噪声、白天短暂阴暗） |
| `drop_partial` / `drop_gap` | true | 去掉首尾不完整夜晚 / 有缺测的夜晚 |
| `min_day_rad_MJ` | 0 | 晴天筛选；0 表示不筛选 |
| `exclude_nights` | [] | 手动排除的夜晚（编号或日期 `YYYY-MM-DD`） |

## 输出

Excel 文件包含：

- `README`：输入文件、时间步长、全部参数（可复现）
- `summary`：长表，每行 = 夜晚 × 样树 × 方法；列 `night_sum`（夜间实测液流总和）、`fit_sum`（夜间充水量）、`remainder = night_sum − fit_sum`（夜间蒸腾量）、`ratio = fit_sum / night_sum`、`breakpoint`、`status`
- `nights`：每个夜晚的起止时间、点数、标记、当日辐射总量与最大值
- `<method>_ratio / _fit_sum / _night_sum`：与 v1 相同结构的宽表（v1 的 `_rate / _night_sum / _sum`）
- `<method>_fitting`（可选）：逐点实测值与拟合值

## 使用前必读

- 方法用于典型晴天。建议用晴天筛选或在“逐夜检查”中排除阴雨天。
- 自动拐点只在夜间数据质量高时使用，否则用固定断点。
- 辐射需要校准；白天低于 5 W m⁻² 或夜间高于 5 W m⁻² 的异常会在“数据检查”中提示。
- 液流数据的突变、异常值需要事先处理。

## 测试

```bash
cd v2 && python3 -m pip install -e ".[test]" && python3 -m pytest -q
```

包括合成数据测试和与 v1 原代码在 `example.xlsx` 上的逐值回归测试。

## 文件结构

```
v2/
├── src/refilltool/
│   ├── methods.py    四种方法
│   ├── nights.py     夜间划分
│   ├── io.py         读取、列识别、数据清理
│   ├── pipeline.py   计算流程与参数
│   ├── export.py     Excel / CSV 输出
│   ├── cli.py        命令行
│   └── app.py        本地图形界面 (Streamlit)
├── tests/
├── start_mac.command
└── start_windows.bat
```
