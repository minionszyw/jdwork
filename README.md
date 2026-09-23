# 表格标准化脚本

`norm.py` 使用 Windows Excel COM 读取 ERP 和店铺后台导出的 `.xlsx`、`.csv` 文件，将原始数据标准化后保存到 `norm` 目录，并保留 Excel 查找和计算公式。

## 运行环境

- Windows
- 已安装 Microsoft Excel
- Python 3.10 或更高版本
- 原始表格放在 `raw` 目录

Excel COM 是脚本读取加密或非标准 XLSX 文件的必要条件。`openpyxl` 不参与表格读写。

## 安装依赖

在 PowerShell 中执行：

```powershell
py -m pip install -r .\requirements.txt
```

如果系统中有多个 Python，建议使用运行脚本的同一个解释器安装：

```powershell
python -m pip install -r .\requirements.txt
```

## 目录结构

```text
.
├── norm.py
├── norm.json
├── requirements.txt
├── raw\       # ERP/店铺后台导出的原始文件
└── norm\      # 脚本生成的标准化文件
```

脚本每次都从 `raw` 重新生成输出，不会把上一次输出再次作为输入处理。

## 使用步骤

1. 从 ERP 或店铺后台导出表格。
2. 按 `norm.json` 中的文件名重命名，并保存到 `raw`。
3. 根据实际工作表名、路径或公式修改 `norm.json`。
4. 关闭正在编辑这些文件的 Excel 窗口。
5. 在项目目录执行：

   ```powershell
   py .\norm.py --config .\norm.json
   ```

6. 在 `norm` 目录查看结果。

运行前只检查配置、文件、工作表和字段，不生成输出：

```powershell
py .\norm.py --check --config .\norm.json
```

## 配置说明

### 路径

```json
{
  "paths": {
    "raw": "raw",
    "norm": "norm"
  }
}
```

路径可以是相对路径或绝对路径。相对路径以 `norm.json` 所在目录为基准。

### 数据源

`sources` 定义 ERP 文件和店铺文件：

- `erp_inventory`：ERP 库存表
- `erp_product`：ERP 商品表
- `erp_ban`：ERP 禁售/控价表
- `erp_combo`：ERP 组合商品表
- `shops`：店铺商品表和销售表

工作表名可以配置；如果配置的工作表不存在，脚本会优先使用 `Sheet1`，再选择第一个非空工作表。

店铺示例：

```json
{
  "name": "百济林",
  "product": "百济林商品.xlsx",
  "product_sheet": "0",
  "sales": "百济林销售.xlsx",
  "sales_sheet": "Sheet1",
  "enabled": true
}
```

将 `enabled` 设置为 `false` 可以跳过某个店铺。

店铺默认使用 `shop_product` 规则。需要不同字段或公式时，在 `rules` 中复制一份规则并修改店铺的 `rule`；删除店铺配置即可停止处理。停用或删除店铺不会自动删除已有的 `norm` 输出文件，脚本会提示旧文件仍然存在。

ERP 数据源可以配置 `output` 和 `reference_range`，店铺销售源可以配置 `sales_reference_range`，避免把工作表列范围写死在 Python 中。

### 字段格式

- `text_columns`：清理空格和制表符后按文本写入，保留前导零。
- `number_columns`：转换为数字，`null`、`--` 等空值会写为空单元格。
- `lookups`：配置 Excel 查找公式。
- `calculations`：配置 Excel 计算公式。
- `number_format`：可选的 Excel 显示格式，例如 `0.0%`。

公式模板支持：

- `{this:字段名}`：当前行字段单元格，例如 `A2`。
- `{range:字段名}`：当前输出表的动态字段范围。
- `{source:别名}`：配置的数据源外部引用。

示例：

```json
{
  "column": "毛利额",
  "formula": "={this:京东价}-{this:SKU进价}"
}
```

脚本不会自动修改公式中的缺失值行为。是否使用 `IFERROR` 由配置的公式模板决定。

## 输出规则

脚本按以下顺序处理：

1. ERP 库存
2. ERP 商品
3. ERP 组合商品
4. 各店铺商品表

销售表只作为店铺商品表的查找源，不会单独生成到 `norm`。

输出文件统一为 `.xlsx`。CSV 输入会通过 Excel COM 打开后另存为 XLSX。

公式写入后，脚本会请求 Excel 重新计算并保存缓存结果，因此打开输出文件时既能看到公式，也能看到计算值。

## 当前标准化字段

### ERP 库存

- `商品代码`：文本

### ERP 商品

- 查找：`可用数量(A - B - C - D)`、`B2C控价金额`、`B2C是否禁售`

### ERP 组合

- 文本：`组合商品代码`、`商品代码`
- 数字：`数量`
- 查找：`进价`、`可用数量(A - B - C - D)`
- 计算：`组合进价`、`组合可用数量`、`单品可拼套数`、`单品进价小计`

### 店铺商品

- 文本：`商家SKU`、`货号`
- 数字：`京东价`、`商品总库存`、`商品可用库存`
- 查找：`进价`、`可用数量(A - B - C - D)`、`B2C控价金额`、`B2C是否禁售`、`SKU进价`、`SKU可用数量`、`商品访客数`、`成交金额`、`成交客户数`
- 计算：`毛利额`、`毛利率`

## 常见问题

### 提示找不到 `win32com`

重新安装依赖：

```powershell
py -m pip install --upgrade -r .\requirements.txt
```

### Excel 进程占用文件

关闭手动打开的 Excel 文件后重试。脚本运行期间不要编辑 `raw` 或 `norm` 中的同名文件。

### 查找结果为 `#N/A`

检查以下内容：

- 查找键两边是否使用了相同的文本格式。
- 原始编码是否包含不可见空格或制表符。
- `norm.json` 中的工作表名是否正确。
- 公式是否需要增加 `IFERROR`。

### 修改了 `raw` 文件名

同步修改 `norm.json` 中对应的 `file` 配置，再重新运行脚本。

### 检查配置失败

先运行 `py .\norm.py --check`。该命令会报告缺少文件、重复店铺、重复输出、未知规则和公式字段错误，但不会修改 `norm`。

## 开发测试

运行不依赖 Excel 的纯逻辑测试：

```powershell
python -m unittest discover -s tests -v
```

## 安全提示

脚本会覆盖 `norm` 目录中同名输出文件，但不会修改 `raw` 原始文件。运行前如需保留旧结果，请先复制 `norm` 目录。
