# Refilling Tool v2

夜间液流组分划分 (nighttime sap flow partitioning) 工具，PWRlab（北京林业大学）。
方法依据：毛白杨夜间液流组分划分方法对比，《植物生态学报》2023.0043（下称“原文”）。

在本机运行，数据不离开电脑；界面只连接 `localhost`。Windows、macOS 均可使用。

## 下载与启动

**方式一：免安装版（推荐给不用 Python 的同学）**

在本仓库 GitHub 页面的 *Releases*（或 *Actions → build-v2* 的运行结果）下载：

- `RefillingTool-Windows-x64.zip`：解压后双击 `RefillingTool.exe`
- `RefillingTool-macOS-AppleSilicon.zip`：解压后双击 `RefillingTool`（首次：右键 → 打开）

浏览器自动打开 `http://localhost:8501`。关闭黑色命令窗口即退出。

**方式二：已装 Python 3.9+**

- macOS：双击 `start_mac.command`
- Windows：双击 `start_windows.bat`

首次运行自动安装依赖（1–3 min；PyPI 连接慢时自动改用清华镜像）。

**方式三：命令行**

```bash
cd v2 && python3 -m pip install -e .
refilltool gui                                   # 图形界面
refilltool run data.xlsx -o results              # 批量计算
refilltool run data.csv --max-rain 0 --area T1_1.3m=0.0123 T1_7.0m=0.0098
refilltool run data.xlsx -c my_config.json       # 用保存的参数复现
refilltool run data.xlsx --legacy                # 复现 v1 结果
```

## 计算方法（与原文对应）

| 步骤 | 原文 | v2 实现 |
|---|---|---|
| 夜间 | 1.3：R < 5 W m⁻² | 阈值可调；短于 30 min 的辐射噪声不切断夜晚；首尾不完整、含缺测的夜晚自动剔除 |
| 液流量 | 公式 (1)：Q = Σ Fd·t·As | 自动检测采样间隔 t；填写边材面积 As 后单位为 mL night⁻¹，否则为 mL m⁻² night⁻¹ |
| 拐点 | 1.5：先平滑，Kneedle（convex, decreasing）检测 P1（快→慢衰减）和 P2（液流基本停止变化） | 默认自动检测；检测失败或早于 30 min 时使用固定值（P1 100 min，P2 300 min）；每个结果记录拐点来源 |
| Re 法 | P2 前为充水，P2 后为蒸腾 | 同原文 |
| Line 法 | 滞后期（约 90 min）结束时液流曲线的切线，切线下方为充水 | 滞后期可调；切线斜率取滞后期内指数拟合在该点的导数；充水不超过实测液流 |
| Exp 法 | P1 前指数拟合 a·e^(−bt)，梯度下降调整 b 使曲线落在液流下方 | 直接求出满足条件的最小 b（梯度下降的精确解，不受 500 步上限影响）；充水不超过实测液流 |
| Et 法 | P2 后液流最小值作为整夜蒸腾速率；蒸腾 = 速率 × 时长，其余为充水 | 同原文（v1 代码此处有误，见下） |
| 高度差预测充水法 | 1.4.3：A = Q1.3m − Q7.0m（公式 2、6），上部液流用 Line 法分为 B 冠层充水、C 夜间蒸腾 | 列名如 `T1_1.3m`、`T1_7.0m` 自动配对，也可手动指定；上部方法可选 |
| 方法评估 | 1.5：C1 与 C2、A1 = (A1+B1) − B2 与 Ah 的 t 检验 | 每棵树及全部树合并输出均值、SE、偏差 %、独立样本与配对 t 检验 P 值，并提示偏差最小的方法 |
| 雨天 | 1.5：只用无降雨日 | 有降雨列时自动剔除当日至夜间结束有降雨的夜晚 |
| 环境因子 | 2.3：夜间均值与组分的相关 | VPD、温度、湿度、土壤水分、风速列自动识别，输出夜间均值，界面可看散点与 Pearson r |

所有方法输出统一：`fit_sum` = 夜间充水 (refilling)，`remainder` = 夜间蒸腾 (transpiration)，`ratio` = 充水占比。

### 与 v1 程序的差别

- **Et 法已按原文修正。** v1 代码没有用 Re 法的 P2（用的是 10 个点），且输出“P2 前实测液流 + 蒸腾基线”，混合了两个组分。示例数据上两者逐夜充水占比的绝对差中位数为 0.32。
- v1 默认使用固定断点（点数），v2 默认按原文自动检测拐点，参数以分钟表示，5、10、30 min 数据通用。
- v1 只对记录直接求和，v2 按原文公式 (1) 乘以采样间隔（及边材面积）。占比 (ratio) 不受影响。
- `--legacy`（界面“v1 兼容模式”）与 v1 程序逐值一致，有自动测试保证。

## 数据格式

一列时间、一列太阳辐射 (W m⁻²)、若干列液流通量密度 Fd（建议 mL m⁻² s⁻¹）。可选：降雨列、环境因子列。列名任意，常见写法自动识别：

| time | solar_radiation | rain_mm | VPD_kPa | T1_1.3m | T1_7.0m |
|---|---|---|---|---|---|
| 2021-07-01 00:00 | 0 | 0 | 0.42 | 0.0213 | 0.0141 |

`examples/synthetic_demo_two_heights.csv` 是**合成数据**（由 `examples/make_synthetic_demo.py` 生成，非实测），用于演示高度差法的输入格式。

## 输出

Excel 文件（同时输出 CSV 和参数 JSON）：

| 工作表 | 内容 |
|---|---|
| README | 输入文件、步长、单位、全部参数 |
| overview | 每个方法 × 每列的充水占比均值 ± SE、充水量、蒸腾量 |
| summary | 长表：夜晚 × 列 × 方法（可直接用于 R/ggplot2） |
| nights | 每夜起止时间、标记、当日辐射、降雨、环境因子夜间均值 |
| height_diff | 高度差法 A、B、C 及占比 |
| method_evaluation | 原文 1.5 的方法评估 |
| `<method>_ratio / _fit_sum / _night_sum` | 与 v1 相同结构的宽表 |
| `<method>_fitting`（可选） | 逐点实测值与充水曲线 |

## 使用建议

- 方法适用于典型晴天。用降雨筛选、晴天筛选，或在“逐夜检查”中排除异常夜晚。
- 在“逐夜检查”里看拐点位置是否合理；不合理时调整平滑窗口或改用固定时长。
- 有两个高度探针时，先看“方法评估”，选偏差最小的方法再做后续分析。
- 辐射需要校准；液流的突变、异常值需要事先处理。

## 开发

```bash
cd v2 && python3 -m pip install -e ".[test]" && python3 -m pytest -q
```

测试包括：与 v1 原代码在 `example.xlsx` 上的逐值回归、各方法的原文公式检验、高度差法对已知茎干充水的还原、单位换算、降雨筛选等。

免安装版由 GitHub Actions（`.github/workflows/build-v2.yml`）在 Windows 和 macOS 上用 PyInstaller 打包；推送 `v2.*` 标签时自动发布到 Releases。

```
v2/
├── src/refilltool/   methods.py 方法 · nights.py 夜间划分 · io.py 读取与列识别
│                     pipeline.py 计算流程 · export.py 输出 · cli.py 命令行 · app.py 界面
├── tests/            自动测试
├── examples/         合成演示数据
├── packaging/        免安装版打包脚本
├── start_mac.command
└── start_windows.bat
```
