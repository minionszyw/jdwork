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
config/             # normalize.json、filter.json
data/raw/           # ERP/店铺导出的原始表（不提交 Git）
data/normalize/     # 标准化结果（不提交 Git）
data/filter/        # 筛选批次（不提交 Git）
src/jdwork/         # CLI 和业务模块
tests/              # 不依赖 Excel 的单元测试
```

`config/` 中的相对路径以各自配置文件所在目录为基准，因此项目目录下的数据路径写为 `../data/raw`、`../data/normalize`、`../data/filter`。

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
   jdw backfill --input .\data\filter\filter-20260923150000.xlsx --dry-run
   jdw backfill --input .\data\filter\filter-20260923150000.xlsx
   ```

回填只允许写入 `config/filter.json` 的 `backfill.fields`，默认按 `店铺 + SKUID` 定位并用 `货号` 校验。写入前会创建 `.bak` 备份。不要在 Excel 中打开正在处理的文件。

## 配置维护

- 在 `sources.shops` 增删店铺；设置 `enabled: false` 可停用。
- 在 `rules` 配置文本、数字、查找和计算字段；公式支持 `{this:字段}`、`{range:字段}`、`{source:别名}`。
- 在 `filter.json.filters` 配置条件和 `eq/ne/lt/lte/gt/gte/in/not_in/contains/is_empty` 等操作符。
- 只把允许人工修改的字段加入 `backfill.fields`，不要加入公式列、`店铺` 或 `类型`。

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
