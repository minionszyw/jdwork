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
config/             # config.json 公共配置、normalize.json、filter.json 流程配置
data/raw/           # ERP/店铺导出的原始表（不提交 Git）
data/normalize/     # 标准化结果（不提交 Git）
data/filter/        # 筛选批次（不提交 Git）
src/jdwork/         # CLI 和业务模块
tests/              # 不依赖 Excel 的单元测试
```

`config/` 中的相对路径以各自配置文件所在目录为基准，因此项目目录下的数据路径写为 `../data/raw`、`../data/normalize`、`../data/filter`。

`config/config.json` 是跨 ERP 和平台通用的配置，使用 `tables` 描述输入表，使用 `rules` 描述格式化和公式动作。每条 table 记录至少包含 `table`、`name`、`type`、`file`；店铺商品表和销售表另外使用相同的 `shop` 值关联。`type` 支持 `erp`、`shop_product`、`shop_sales`。`normalize.json` 只配置路径、Excel 专用信息和每张表的 sheet/输出文件；`filter.json` 只配置筛选和回填。

## 闭环使用

1. 按 `config/normalize.json` 重命名 ERP/店铺导出的 `.xlsx` 或 `.csv`，放入 `data/raw/`。
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

回填只允许写入 `config/filter.json` 的 `backfill.fields`，默认按 `店铺 + SKUID` 定位并用稳定字段校验。写入前会创建 `.bak` 备份。不要在 Excel 中打开正在处理的文件。

同一店铺和 SKUID 命中多个规则时，回填会合并各行的字段修改：不同字段可以同时回填；同一字段的不同修改值会报告冲突并停止该批次。回填后应重新执行 `normalize` 和 `filter`，下一轮使用新的筛选批次。

## 配置维护

- 在 `config.json` 的 `tables` 增删 ERP 或店铺表；店铺商品表和销售表设置相同的 `shop`。
- 在 `config.json` 的 `rules` 配置 `format`（`text`/`number`）和 `function`（Excel 公式）动作；公式支持 `{this:字段}`、`{range:字段}`、`{source:别名}`。
- 在 `filter.json.filters` 配置条件和 `eq/ne/lt/lte/gt/gte/in/not_in/contains/is_empty` 等操作符。
- 只把允许人工修改的字段加入 `backfill.fields`，不要加入公式列、`店铺`、`类型` 或 `SKUID`。如果要修改 `货号`，请把不会被编辑的稳定字段（例如 `商品编码`）配置到 `backfill.verify_fields`。
- 新增指标或筛选规则只需修改 JSON；Python 模块提供通用格式化、公式、筛选和回填能力。
- `missing_encoding` 可筛选 `商家SKU` 或 `货号` 为 `--` 的商品。修改筛选表中的编码后运行回填即可，无需重新导出原始表；回填后重新运行 `jdw normalize` 刷新标准化结果。

显式配置路径会覆盖默认值：

```powershell
jdw normalize --config .\config\normalize.json
jdw filter --config .\config\filter.json
```

## 开发验证

```powershell
python -m unittest discover -s tests -v
python -m compileall -q src\jdwork
```

真实 Excel 验证建议依次执行 `jdw normalize --check`、`jdw filter --check` 和 `jdw backfill --dry-run`。不要提交 `data/raw/`、`data/normalize/`、`data/filter/`、`docs/` 或客户数据。
