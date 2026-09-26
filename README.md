# jdwork

`jdwork` 将 ERP 和店铺后台导出的 Excel/CSV 表格标准化、筛选并回填。表格通过 Windows Excel COM 读取，因此支持加密或格式不规范的工作簿，并保留标准化输出中的 Excel 公式。

## 环境与安装

- Windows
- Microsoft Excel
- Python 3.10+

在仓库根目录执行本地安装（也可使用 `pipx install .`）：

```powershell
py -m pip install .
```

开发时使用可编辑安装：

```powershell
py -m pip install -e .
```

## 目录

```text
config/             # 公共路径/表配置、标准化/筛选/回填专用配置
data/raw/           # ERP/店铺导出的原始表（不提交 Git）
data/normalize/     # 标准化结果（不提交 Git）
data/filter/        # 筛选批次（不提交 Git）
src/jdwork/         # CLI 和业务模块
tests/              # 不依赖 Excel 的单元测试
```

## 配置维护

所有配置文件使用展开缩进，路径均相对于 `config/` 所在目录。

| 文件 | 内容 |
| --- | --- |
| `config.json` | 公共 `paths`、`tables` 和默认 `sheet`；表记录包含 `table`、`name`、`type`、`file`、`shop`。 |
| `normalize.json` | 按表分组的标准化规则；每个 `column` 使用 `format`（`text`/`number`）或 `function`（Excel 公式）。 |
| `filter.json` | `filters`、条件及 `eq/ne/lt/lte/gt/gte/in/not_in/contains/is_empty` 等操作符。 |
| `backfill.json` | 可回填的 `fields` 及 `verify_fields`、备份等安全设置。 |

维护配置时，在 `config.json` 的 `tables` 中增删 ERP 或店铺表，并为同一店铺的商品表和销售表设置相同的 `shop`。标准化规则按 `table` 分组，公式支持 `{this:字段}`、`{range:字段}` 和 `{source:别名}`；`shop_product` 规则可应用到所有店铺商品表。仅将允许人工修改的字段加入 `backfill.json` 的 `fields`，不要加入公式列、`店铺`、`类型` 或 `SKUID`；修改 `货号` 时，应配置不会被编辑的稳定字段（如 `商品编码`）作为 `verify_fields`。

## 闭环使用

1. 按 `config/config.json` 的 `tables` 重命名 ERP/店铺导出的 `.xlsx` 或 `.csv`，放入 `data/raw/`。
2. 标准化：

   ```powershell
   jdw normalize
   jdw normalize --check
   ```

3. 按 `config/filter.json` 筛选：

   ```powershell
   jdw filter
   jdw filter --check
   ```

   输出为 `data/filter/filter-{batch_id}.xlsx`，可用 `--batch-id 20260923150000` 指定批次号。
4. 人工修改筛选表后先预览回填，再应用：

   ```powershell
   jdw backfill --dry-run
   jdw backfill
   ```

   省略 `--input` 时，命令会在 `data/filter/` 中按文件名批次号选择最大的
   `filter-{14 位 batch_id}.xlsx`。也可以显式指定某个筛选文件。

回填只允许写入 `config/backfill.json` 的 `fields`，默认按 `店铺 + SKUID` 定位并用稳定字段校验。写入前会创建 `.bak` 备份。不要在 Excel 中打开正在处理的文件。

同一店铺和 SKUID 命中多个规则时，回填会合并各行的字段修改：不同字段可以同时回填；同一字段的不同修改值会报告冲突并停止该批次。回填后应重新执行 `normalize` 和 `filter`，下一轮使用新的筛选批次。

## 开发验证

```powershell
python -m unittest discover -s tests -v
python -m compileall -q src\jdwork
```

真实 Excel 验证建议依次执行 `jdw normalize --check`、`jdw filter --check` 和 `jdw backfill --dry-run`。不要提交 `data/raw/`、`data/normalize/`、`data/filter/`、`docs/` 或客户数据。
