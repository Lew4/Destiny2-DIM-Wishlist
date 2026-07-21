# Destiny 2 DIM 双输入愿望单生成器

把中文武器推荐转换成 [Destiny Item Manager](https://app.destinyitemmanager.com/) 可导入的 Wishlist。一个项目支持两种来源：

- `text`：武器名和 perk 都是 CSV/XLSX 单元格文字。
- `icon`：武器名是文字，但 PVE/PVP 三、四号位 perk 是 XLSX 内嵌图片。

两种模式共用同一套 Bungie Manifest、武器版本检索和 uint32 hash 处理逻辑。最终 perk hash 始终由具体武器版本的合法 socket 决定。

## 项目结构

```text
dim_wishlist_builder.py       统一兼容入口
dim_wishlist/
  app.py                      text/icon 命令分发
  cli.py                      文字表工作流
  table.py                    文字CSV/XLSX解析
  wishlist.py                 文字推荐的版本兼容与DIM生成
  reports.py                  文字推荐审计报告
  icon_cli.py                 图标XLSX工作流
  icon_xlsx.py                drawing anchor与内嵌图片提取
  icon_matching.py            图像归一化和官方图标识别
  icon_wishlist.py            trait socket校验和DIM生成
  icon_reports.py             图标匹配审核报告
  manifest.py                 共享Manifest索引与socket能力
  config.py/icon_config.py    两种工作流配置
  models.py/icon_models.py    数据模型
  utils.py                    共享工具
input/
  命运2-...推荐-Sheet1.csv     文字推荐表
  d2.xlsx                     图标型XLSX
  world_sql_content_*.content 本地Manifest（不提交）
tests/                        单元和提取回归测试
outputs/                      两种模式的统一输出目录
```

## 安装和Manifest

需要 Python 3.9 或更高版本：

```bash
python3 -m pip install -r requirements.txt
```

依赖包括 `openpyxl`、`Pillow` 和 `numpy`。从 Bungie 下载对应语言的 `world_sql_content_*.content` 放在 `input/`。Manifest 通常有数百 MB，已被 `.gitignore` 排除。

## 统一命令

```bash
python3 dim_wishlist_builder.py --help
python3 dim_wishlist_builder.py text --help
python3 dim_wishlist_builder.py icon --help
```

也支持：

```bash
python3 -m dim_wishlist text
python3 -m dim_wishlist icon
```

没有写子命令时默认使用 `text`，因此原来的运行方式完全兼容：

```bash
python3 dim_wishlist_builder.py
```

## 文字表模式

```bash
python3 dim_wishlist_builder.py text \
  --input input/命运2-凯旋丰碑全种类武器推荐-Sheet1.csv \
  --manifest input/world_sql_content_xxx.content \
  --output-dir outputs
```

支持 `.csv`、`.xlsx` 和 `.xlsm`。工具会识别 `名字`、`武器`、`name` 等武器列以及所有包含 `Perk` 的列，保留重复表头，跳过分段中重复出现的表头，并把多行或 `/`、`、`、`,` 分隔的 perk 展开为笛卡尔积。

四个输入列分别代表枪管/瞄具、弹匣/电池、第一特性和第二特性。实际匹配会根据名称覆盖率推断它们对应的真实 socket，不盲目依赖列序。

默认处理同名武器的全部历史版本。旧版本缺少部分推荐 perk 时会生成 `[兼容子集]`；使用 `--version-perk-policy strict` 可改为整版跳过。

文字模式输出：

- `dim_wishlist_resolved.txt`

默认只保留最终 Wishlist。需要排查时加 `--diagnostics`，才会额外生成
`dim_wishlist_unresolved.csv`、审计、候选和提取报告。

## 图标XLSX模式

首次建议只检查XLSX结构，不读Manifest、不访问网络：

```bash
python3 dim_wishlist_builder.py icon \
  --input input/d2.xlsx \
  --output-dir outputs \
  --run-mode extract_only
```

`extract_only` 本身属于诊断模式，会保留提取报告和图标。

完整生成：

```bash
python3 dim_wishlist_builder.py icon \
  --input input/d2.xlsx \
  --manifest input/world_sql_content_xxx.content \
  --output-dir outputs \
  --run-mode full
```

### 它如何识别图标

这不是OCR，也不依赖屏幕截图坐标。工具直接打开XLSX压缩结构，读取 worksheet、drawing relationship 和 two-cell/one-cell anchor，将每张内嵌图片定位到：

```text
武器名称 → PVE/PVP → trait_3/trait_4 → 槽内第几个候选
```

之后执行两阶段解析：

1. 每个唯一Excel图标只在全局官方普通特性图标库中识别一次。
2. 对每个武器版本，只在严格的普通 trait socket 中选择该视觉对应的实际 hash。

每把武器推荐行下一行的 PVE、PVP 说明会分别写入对应武器区块的
`//notes:`。左侧 B–I/PVE 和 K–R/PVP 的 perk 图例说明只用于阅读，不会写入
愿望单。

如果图标在工作簿中误放到另一列，但在该武器版本中只对应一个真实 trait
socket，程序会按 Manifest 自动纠正到实际三号位或四号位。工作簿中同名但武器
类型不同的误译也可以用“武器名 + 武器类型”覆盖，避免例如烈日速射斥候
“受托”被误匹配成手炮“信任”。

同一中文名在工作簿中明确代表不同来源版本时，可以按“Excel 行号 + 武器名”
限定 hash；明确取消的推荐会记录为排除项，不参与规则组合，也不会出现在错误报告中。
普通同名历史版本仍会参与兼容性检查；正式愿望单会保留所有完整支持该组推荐的
正式历史版本，但部分兼容版本和指向其他正式 hash 的内部派生定义只归入历史诊断。
明确按来源拆分的版本（例如救赎花园版与万神殿版“鲁莽神谕”）也会分别保留。

少数工作簿会把非trait perk图标放在trait展示区域。此类图标使用按内容SHA-256记录的“名称 + 语义槽位”覆盖，仍必须通过具体武器socket校验。例如维卡拉微冲4的 `AG226` 实际是二号位“超频散热器”；工具会在socket 2选择普通版hash，而不会误当成三号特性或选择强化版。

官方图标首次运行会从 Bungie 下载并缓存在 `outputs/.official_icon_cache/`。匹配依次利用原文件SHA-256、透明边缘裁切后的规范像素、透明轮廓SSIM/IoU、灰度结构、边缘余弦、dHash以及±2像素平移搜索。默认阈值为相似度 `0.935`、不同perk名称候选间距 `0.025`。

图标模式输出：

- `dim_icon_wishlist_resolved.txt`：最终PVE/PVP Wishlist。

默认只保留最终 Wishlist 和隐藏的官方图标缓存。需要审核时加 `--diagnostics`，
才会额外生成以下文件：

- `icon_global_review.html`：184个唯一图标的人工审核页。
- `icon_global_matches.csv` / `icon_global_unresolved.csv`：全局视觉识别。
- `icon_matches.csv`：全部武器版本和 socket 的详细解析。
- `icon_unresolved.csv` / `icon_source_issues.csv`：跨全部版本后仍然存在的真实问题。
- `icon_excluded_recommendations.csv`：人工明确取消、不参与生成的推荐。
- `icon_history_compatibility.csv` / `icon_history_summary.csv`：旧版本缺少perk的兼容记录；只要至少一个版本完整匹配，就不会算作错误。
- `dim_icon_wishlist_audit.csv`：最终组合审核。
- `icon_weapon_candidates.csv`：同名武器候选。
- `icon_extracted.csv`、`extracted_icons/`：XLSX提取结果。

诊断运行后可查看 `icon_global_review.html`，再处理 `icon_unresolved.csv` 中的问题。

## 输出和版本控制

两种模式统一写入 `outputs/`，文件名互不冲突。普通运行时该目录只显示最终
两个 Wishlist 文本；审计 CSV、HTML 和提取图标仅在 `--diagnostics` 模式生成。
隐藏的 `.official_icon_cache/` 会保留以避免重复下载。仓库只提交最终两个 Wishlist。

## 测试

```bash
python3 -m unittest discover -s tests -v
```

测试覆盖文字表解析、signed/unsigned hash、旧版本兼容降级、真实样例XLSX的drawing提取、图标平移容忍、错列自动归位、同名武器类型纠正和特殊二号位普通/强化perk选择。当前完整真实Manifest回归基线为：4120个drawing、3221个perk图标位置、184/184个唯一perk图标识别成功；完整正式版本经DIM语义去重后生成Wishlist规则，部分兼容版本仅归入历史诊断。
