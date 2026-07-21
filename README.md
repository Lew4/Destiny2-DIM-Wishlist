# Destiny 2 DIM 中文愿望单生成器

把中文武器推荐表转换成 [Destiny Item Manager](https://app.destinyitemmanager.com/) 可导入的 Wishlist。工具会使用本地 Bungie Manifest，在具体武器版本的具体 socket 内解析 perk，避免同名 perk 或历史武器版本使用错误 hash。

## 项目结构

```text
dim_wishlist_builder.py       兼容入口，仍可直接运行
dim_wishlist/
  cli.py                      命令行和完整流程编排
  config.py                   默认配置、生成策略和公共常量
  manifest.py                 Manifest 读取、socket 推断和 perk 解析
  table.py                    CSV/XLSX 读取、列识别和组合展开
  wishlist.py                 武器版本处理及 DIM 规则生成
  reports.py                  审计和候选报告
  models.py                   数据模型
  utils.py                    文本、hash、文件和 CSV 工具
tests/                        不依赖真实 Manifest 的单元测试
examples/                     示例推荐表
outputs/                      默认输出目录
```

## 安装

需要 Python 3.9 或更高版本。CSV 输入只使用标准库；读取 XLSX 需要 `openpyxl`：

```bash
python3 -m pip install -r requirements.txt
```

从 Bungie 下载对应语言的 `world_sql_content_*.content`，放到项目根目录。Manifest 通常有数百 MB，已被 `.gitignore` 排除，不会提交到仓库。

## 运行

仓库中的文件名已经设为默认值，直接运行即可：

```bash
python3 dim_wishlist_builder.py
```

也可以显式指定路径和策略：

```bash
python3 dim_wishlist_builder.py \
  --input examples/sample_recommendations.csv \
  --manifest /path/to/world_sql_content.content \
  --output-dir outputs \
  --weapon-version-mode all \
  --version-perk-policy drop_unsupported
```

查看完整参数：

```bash
python3 dim_wishlist_builder.py --help
```

也支持包入口：

```bash
python3 -m dim_wishlist
```

## 输入格式

支持 `.csv`、`.xlsx` 和 `.xlsm`。武器列默认识别 `名字`、`武器`、`name` 等名称；所有列名包含 `Perk` 的列会作为词条列。重复的 `Perk` 表头会被保留，多行或以 `/`、`、`、`,` 等分隔的 perk 会自动展开为组合。

仓库原始表按武器类型分段，每段重复表头。解析器会识别第一段真实表头并跳过后续重复表头。

四个 perk 列依次对应：

1. 枪管、瞄具、弓弦等第一槽位
2. 弹匣、电池、箭杆等第二槽位
3. 第一特性槽
4. 第二特性槽

实际生成时不会盲目相信列序，而是根据推荐词条在各 socket 中的名称覆盖率推断对应关系，再以 socket 分类和原始顺序打破平局。

## 多版本兼容策略

默认设置为：

```text
weapon_version_mode = all
version_perk_policy = drop_unsupported
```

因此同名历史版本会分别生成：

- 当前版本支持全部推荐 perk：生成完整笛卡尔积。
- 只支持部分 perk：生成兼容子集，并在标题标记 `[兼容子集]`。
- 某个槽位没有兼容 perk：省略该槽位，继续匹配其他槽位。
- 一个推荐 perk 都不支持：跳过该版本。

使用 `--version-perk-policy strict` 可以在任一推荐 perk 不受支持时跳过整个版本；使用 `--weapon-version-mode single` 可以只选择稳定排序后的第一个同名版本。

## 输出

默认写入 `outputs/`：

- `dim_wishlist_resolved.txt`：可导入 DIM 的愿望单。
- `dim_wishlist_unresolved.csv`：不兼容、未匹配或完全跳过的项目。
- `dim_wishlist_resolved_audit.csv`：每条 DIM 规则实际采用的版本、槽位和 perk。
- `dim_wishlist_perk_candidates.csv`：每个版本/socket 的匹配过程。
- `dim_wishlist_weapon_candidates.csv`：同名武器版本候选。
- `dim_wishlist_extracted.csv`：原始表解析和笛卡尔积展开结果。

审计 CSV 是可重复生成的本地文件，默认不提交；仓库保留最终的 `outputs/dim_wishlist_resolved.txt` 供 DIM 使用。

## 测试

```bash
python3 -m unittest discover -s tests -v
```

测试覆盖 hash 的 signed/unsigned 转换、重复表头和多选 perk 展开，以及同名武器旧版本的兼容降级。
