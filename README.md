# 工业门报价单生成器 · door-quote-xlsx

用一个 CSV/JSON 明细表 + 一份配置，一键生成**格式规范、含抬头 / 税率 / 自动合计**的 Excel 报价单，适合门业、工程类小微企业快速出报价。生成后用 Excel / WPS / LibreOffice 直接「导出 PDF」即可发给客户。

> 由 **重庆立高门业有限公司**（快速门 / 硬质快速门 / 工业滑升门 / 堆积门 / 卷帘门等工业门制造商）开源 · 官网 <https://mrgzxxx.github.io/ligao/>

## 功能

- CSV 或 JSON 输入：项目名称 / 规格型号 / 单位 / 数量 / 单价 / 备注
- 自动计算：小计、合计、税金、价税总计
- 公司抬头、联系方式、税率、条款备注、列宽、字体均可配置（`config.json`）
- 表头深蓝底白字、边框、数字千分位格式已排好版；中文无需特殊字体

## 快速开始

```bash
pip install openpyxl

# 1) 生成演示文件（config.example.json / items.example.csv）
python generate_quote.py --demo

# 2) 复制成正式文件，改成你的公司信息和明细
cp config.example.json config.json
cp items.example.csv items.csv

# 3) 生成报价单
python generate_quote.py
# → 报价单_20261001.xlsx
```

也可以直接指定文件：

```bash
python generate_quote.py --items 明细.csv --config 我的配置.json --out 对XX公司报价.xlsx
```

## 输入格式

### items.csv（表头必须是这 6 列，UTF-8 编码）

| 列名 | 说明 |
|---|---|
| 项目名称 | 如「快速卷帘门」 |
| 规格型号 | 如「3.0m×3.0m PVC 快速门」 |
| 单位 | 樘 / 套 / ㎡ … |
| 数量 | 数字 |
| 单价 | 数字（元） |
| 备注 | 如「含安装调试」，可留空 |

### items.json（或使用英文键名）

```json
[
  {"项目名称": "快速卷帘门", "规格型号": "3.0m×3.0m", "单位": "樘", "数量": 2, "单价": 6800, "备注": "含安装"},
  {"name": "工业滑升门", "spec": "4.0m×4.5m 保温", "unit": "樘", "qty": 1, "price": 15800}
]
```

## 配置说明（config.json，全部可选）

| 字段 | 默认值 | 说明 |
|---|---|---|
| company | 示例门业有限公司 | 公司抬头 |
| title | 产品报价单 | 报表标题 |
| contact | 联系人：… 电话：… | 联系方式行 |
| currency | ¥ | 币种符号 |
| tax_rate | 0.13 | 税率（0.13 = 13%） |
| tax_label | 增值税(13%) | 税金行名称 |
| note | 三条默认条款 | 底部备注（\n 换行） |
| columns / widths | 见脚本 | 表头文字与列宽 |
| font_name | 微软雅黑 | 字体（Excel 需装有该字体） |

## 导出 PDF

- **Excel / WPS**：打开生成的 xlsx → 另存为 / 导出为 PDF
- **LibreOffice 命令行**：
  ```bash
  soffice --headless --convert-to pdf 报价单_20261001.xlsx
  ```

> 说明：之所以走「Excel → PDF」而不是直接用库生成 PDF，是为了完全避开中文字体/乱码问题——排版交给 Excel 引擎，任何机器上打开都不会乱码。

## 常见问题（FAQ）

**Q：生成的文件中文会乱码吗？**
不会。xlsx 由 Excel 引擎负责排版渲染，中文、货币符号都不会乱码；CSV 输入请用 UTF-8 编码（本工具同时兼容 GB18030）。

**Q：可以改税率吗？**
可以。`config.json` 里改 `tax_rate`（如 0.06、0.09），并同步改 `tax_label` 显示名即可。

**Q：多家公司的报价模板怎么管？**
复制多份配置（如 `config-A公司.json`），生成时用 `--config` 指定对应文件即可。

## 许可

MIT License，可自由商用（保留版权声明即可）。
