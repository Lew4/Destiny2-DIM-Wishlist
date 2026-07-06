# DIM 中文推荐表转愿望单工具

用途：把中文武器推荐表转换成 Destiny Item Manager 可导入的 wish list 文本。

输入：

1. 中文推荐表：`.xlsx` 或 `.csv`
2. Bungie manifest 数据库文件：你下载的 `world_sql_content_xxx.content`

输出：

- `dim_wishlist_resolved.txt`：DIM 可导入愿望单
- `dim_wishlist_unresolved.csv`：没有匹配成功的武器 / perk
- `dim_wishlist_extracted.csv`：从原始推荐表识别并展开后的内容，用来检查是否漏读、错拆分

## 1. 安装依赖

```bash
cd dim_wishlist_builder
python3 -m pip install -r requirements.txt
```

## 2. 准备 Bungie manifest

你已经下载的文件类似：

```text
world_sql_content_22b6eb96bbcaa631746b584b52bcc2a6.content
```

可以直接传给脚本。脚本会自动判断它是压缩包还是 SQLite 数据库；如果是压缩包，会自动解压。

## 3. 运行

```bash
python3 dim_wishlist_builder.py \
  --input /path/to/命运2-凯旋丰碑全种类武器推荐.xlsx \
  --manifest /path/to/world_sql_content_xxx.content \
  --out dim_wishlist_resolved.txt \
  --unresolved dim_wishlist_unresolved.csv \
  --extracted dim_wishlist_extracted.csv
```

## 4. 输入表格要求

推荐使用以下列名：

```text
名字, Perk, Perk, Perk 1, Perk 2, 注释, Tier
```

说明：

- `名字`：武器名。
- 所有包含 `Perk` 的列都会作为 perk 列。
- 单元格中多行 perk 会自动展开。
- 例如 `滑射\n快速命中` 会被视为两个可选 perk，并与其他列做组合。
- `注释`、`Tier`、`Rank` 会写入分组级 `//notes:`，格式接近官方工具导出的 block note。



## 5. 先检查识别结果

脚本会额外输出：

```text
dim_wishlist_extracted.csv
```

这个文件不依赖 hash 匹配，只反映脚本从原始推荐表里读到了什么。重点看这些列：

- `source_row`：原始表中的数据行编号；
- `weapon`：识别出的武器名；
- `expanded_perks`：展开后的单条 perk 组合；
- `notes`：由 `Tier`、`Rank`、`注释` 合成的说明；
- `raw_perk_columns_json`：原始 perk 单元格内容；
- `parsed_perk_columns_json`：每个 perk 列被拆分成了哪些候选 perk。

如果这里已经漏了武器、perk 被错误拆开，或者多行内容没有展开对，先改原始表或分隔符，再重新生成愿望单。

## 6. 输出格式

当前脚本输出接近官方工具样式：

```text
// 武器名 - recommended
//notes: Tier S | PvE
dimwishlist:item=武器hash&perks=perk1hash,perk2hash,perk3hash
dimwishlist:item=另一个版本hash&perks=perk1hash,perk2hash,perk3hash
```

说明：

- `// 武器名 - recommended` 是普通注释。
- `//notes:` 是 block note，会作用于下面连续的 DIM 规则。
- 每条 `dimwishlist:` 规则末尾不再追加 `#notes:`。

## 7. 同名不同版本

如果 manifest 中同一个武器名对应多个武器 hash，脚本会为每个版本都生成同一套 perk 组合。

例如：

```text
玫瑰 + 滑射 + 首发
```

如果查到 3 个 `玫瑰` 版本，会生成 3 条 DIM 规则。

## 8. 未匹配结果

如果武器或 perk 没查到，会写入：

```text
dim_wishlist_unresolved.csv
```

你需要重点看这个文件。常见原因：

- 表格里的中文名和 manifest 名字不完全一致；
- perk 是俗称，不是官方译名；
- 写了组合说明，而不是单个 perk 名；
- manifest 语言版本不匹配。

## 9. 导入 DIM

把生成的 `dim_wishlist_resolved.txt` 放到 GitHub public gist 或 public repo，取 raw 链接，在 DIM 设置里添加为外部 wish list。

也可以先打开文件检查格式。
