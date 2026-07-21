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
examples/
  sample_recommendations.csv  文字表示例
  d2.xlsx                     图标型XLSX验证样例
tests/                        单元和提取回归测试
outputs/                      两种模式的统一输出目录
```

## 安装和Manifest

需要 Python 3.9 或更高版本：

```bash
python3 -m pip install -r requirements.txt
```

依赖包括 `openpyxl`、`Pillow` 和 `numpy`。从 Bungie 下载对应语言的 `world_sql_content_*.content` 放在项目根目录。Manifest 通常有数百 MB，已被 `.gitignore` 排除。

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
  --input 命运2-凯旋丰碑全种类武器推荐-Sheet1.csv \
  --manifest world_sql_content_xxx.content \
  --output-dir outputs
```

支持 `.csv`、`.xlsx` 和 `.xlsm`。工具会识别 `名字`、`武器`、`name` 等武器列以及所有包含 `Perk` 的列，保留重复表头，跳过分段中重复出现的表头，并把多行或 `/`、`、`、`,` 分隔的 perk 展开为笛卡尔积。

四个输入列分别代表枪管/瞄具、弹匣/电池、第一特性和第二特性。实际匹配会根据名称覆盖率推断它们对应的真实 socket，不盲目依赖列序。

默认处理同名武器的全部历史版本。旧版本缺少部分推荐 perk 时会生成 `[兼容子集]`；使用 `--version-perk-policy strict` 可改为整版跳过。

文字模式输出：

- `dim_wishlist_resolved.txt`
- `dim_wishlist_unresolved.csv`
- `dim_wishlist_resolved_audit.csv`
- `dim_wishlist_perk_candidates.csv`
- `dim_wishlist_weapon_candidates.csv`
- `dim_wishlist_extracted.csv`

## 图标XLSX模式

首次建议只检查XLSX结构，不读Manifest、不访问网络：

```bash
python3 dim_wishlist_builder.py icon \
  --input examples/d2.xlsx \
  --output-dir outputs \
  --run-mode extract_only
```

完整生成：

```bash
python3 dim_wishlist_builder.py icon \
  --input examples/d2.xlsx \
  --manifest world_sql_content_xxx.content \
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

官方图标首次运行会从 Bungie 下载并缓存在 `outputs/.official_icon_cache/`。匹配依次利用原文件SHA-256、透明边缘裁切后的规范像素、透明轮廓SSIM/IoU、灰度结构、边缘余弦、dHash以及±2像素平移搜索。默认阈值为相似度 `0.935`、不同perk名称候选间距 `0.025`。

图标模式输出：

- `dim_icon_wishlist_resolved.txt`：最终PVE/PVP Wishlist。
- `icon_global_review.html`：184个唯一图标的人工审核页。
- `icon_global_matches.csv` / `icon_global_unresolved.csv`：全局视觉识别。
- `icon_matches.csv` / `icon_unresolved.csv`：武器版本和socket解析。
- `dim_icon_wishlist_audit.csv`：最终组合审核。
- `icon_weapon_candidates.csv`：同名武器候选。
- `icon_extracted.csv`、`extracted_icons/`：XLSX提取结果。

建议先查看 `icon_global_review.html`，再处理 `icon_unresolved.csv` 中的版本不兼容项。

## 输出和版本控制

两种模式统一写入 `outputs/`，文件名互不冲突。审计CSV、HTML、提取图标和官方图标缓存均可重复生成，默认不提交；仓库只保留最终两个Wishlist文本。

## 测试

```bash
python3 -m unittest discover -s tests -v
```

测试覆盖文字表解析、signed/unsigned hash、旧版本兼容降级、真实样例XLSX的drawing提取，以及图标缩放和平移容忍。完整真实Manifest回归应保持：4115个drawing、3216个perk图标位置、184个唯一perk图标和6105条图标Wishlist规则。
