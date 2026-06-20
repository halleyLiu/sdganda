# SDGanda 机体数据导出

这个目录包含一个导出脚本，以及脚本从 SDGanda 接口抓取后生成的数据文件。

## 文件说明

- `fetch_sdganda_units.py`：抓取所有机体的 `units` 信息，并逐个抓取近 N 天胜率数据。
- `YYYYMMDD/units.json`：`/units` 接口返回的原始机体列表。
- `YYYYMMDD/win_rates.json`：按机体 ID 保存的原始胜率数据。
- `YYYYMMDD/combined_units_win_rate.json`：把机体信息和胜率数据合并后的扁平 JSON。
- `YYYYMMDD/sdganda_units_YYYYMMDD.xlsx`：用于筛选、排序、查看的 Excel 文件。

例如在 2026-06-20 执行，输出目录就是 `20260620/`。

## 如何执行

在 PowerShell 中运行：

```powershell
& 'C:\Users\halley\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' 'C:\Users\halley\Documents\Codex\2026-06-20\https-www-sdganda-com-units-15298\outputs\fetch_sdganda_units.py'
```

脚本会自动创建以当天日期命名的目录，格式为 `YYYYMMDD`。

常用参数：

```powershell
# 只抓前 3 台机体，用于测试
& 'C:\Users\halley\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' '.\fetch_sdganda_units.py' --sample 3

# 指定输出日期目录
& 'C:\Users\halley\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' '.\fetch_sdganda_units.py' --date 20260620

# 指定胜率统计窗口，默认是 14 天
& 'C:\Users\halley\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' '.\fetch_sdganda_units.py' --days 14
```

## 数据来源

脚本使用以下接口：

- 机体列表：`https://dbapi.sdganda.com/api/units?limit=1000`
- 胜率数据：`https://dbapi.sdganda.com/api/sdgore/win-rate/{unit_id}/{days}`

默认 `{days}` 为 `14`，也就是近 14 天胜率。

## Excel 表格说明

生成的 Excel 只有一个工作表，名称为 `units_win_rate`。

第一行已冻结，并开启筛选。主要字段会排在最前面：

- `ID`：机体编号。
- `Name`：中文机体名。
- `EngName`：接口返回的英文机体名。
- `品阶`：已转换成可读文本的品阶，例如 `S`、`SS`、`SR`、`SU`。
- `Rank`：接口返回的原始品阶数字。
- `RankType`：接口返回的原始品阶类型数字。
- `Type`：接口返回的机体类型数字。
- `win_times`：统计窗口内胜利次数。
- `lose_times`：统计窗口内失败次数。
- `draw_times`：统计窗口内平局次数。
- `win_rate`：胜率，小数形式保存，Excel 中按百分比格式显示。
- `lose_rate`：败率，小数形式保存，Excel 中按百分比格式显示。
- `averageRating`：站内用户评分均值，如果接口中存在该字段。
- `totalRatings`：站内用户评分数量，如果接口中存在该字段。

这些主要字段之后，脚本还会保留 `/units` 接口中的其他标量字段，例如：

- `Hp`
- `Attack`
- `Defence`
- `MoveLevel`
- `DashLevel`
- `WeaponID1`、`WeaponID2` 等武器 ID 字段
- `SkillID1`、`SkillID2` 等技能 ID 字段

完整的 `skills`、`weapons`、`rating` 等嵌套对象或数组会保留在 `units.json` 中，但不会展开到 Excel。这样可以让 Excel 主表保持扁平结构，方便筛选和排序。

## 品阶显示逻辑

接口里与品阶有关的字段有两个：

- `Rank`：基础品阶数字。
- `RankType`：品阶类型数字。

脚本先根据 `Rank` 得到基础品阶：

| Rank | 基础显示 |
| ---: | --- |
| 0 | C |
| 1 | C |
| 2 | B |
| 3 | A |
| 4 | S |

然后根据 `RankType` 决定是否追加后缀：

| RankType | 追加后缀 |
| ---: | --- |
| 3 | S |
| 4 | R |
| 5 | U |

其他 `RankType` 不追加后缀。

示例：

| Rank | RankType | 最终 `品阶` |
| ---: | ---: | --- |
| 4 | 2 | S |
| 4 | 3 | SS |
| 4 | 4 | SR |
| 4 | 5 | SU |
| 3 | 2 | A |
| 2 | 2 | B |

也就是说：

- `Rank=4, RankType=2` 显示为 `S`
- `Rank=4, RankType=3` 显示为 `SS`
- `Rank=4, RankType=4` 显示为 `SR`
- `Rank=4, RankType=5` 显示为 `SU`

这个逻辑与 SDGanda 页面前端显示逻辑一致：页面先显示基础品阶，只在 `RankType` 为 `3`、`4`、`5` 时分别追加 `S`、`R`、`U`。

## 输出数据说明

`win_rates.json` 按机体 ID 保存胜率数据，例如：

```json
{
  "15298": {
    "unit_id": 15298,
    "win_times": 5162,
    "lose_times": 4899,
    "draw_times": 33,
    "win_rate": 0.5113929,
    "lose_rate": 0.48533782
  }
}
```

字段含义：

- `unit_id`：机体编号。
- `win_times`：胜利次数。
- `lose_times`：失败次数。
- `draw_times`：平局次数。
- `win_rate`：胜率，小数形式，例如 `0.5113929` 表示约 `51.14%`。
- `lose_rate`：败率，小数形式。

`combined_units_win_rate.json` 是合并后的逐机体数据，也是生成 Excel 的数据来源。

如果某个机体的胜率接口请求失败，脚本仍会保留该机体行，并把：

- `win_times`
- `lose_times`
- `draw_times`
- `win_rate`
- `lose_rate`

这些字段留空，同时在 `win_rate_error` 中记录失败信息。
