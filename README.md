# CO 外贸跟单工具

面向服装外贸跟单场景的 Windows 桌面工具，将订单整理、港口与数量汇总、验货资料、装箱单、出货表和箱唛处理集中到一个中文界面，减少重复录入和手工整理表格的工作。

工具基于 Python 与 Tkinter 开发，支持源码运行，也可打包为 Windows EXE。当前版本为 **1.1.3**。

## 功能一览

| 功能模块 | 主要用途 |
| --- | --- |
| ① 订单明细导入 | 读取 PO CSV，配置颜色、厂款号、工厂出厂期和目的港映射，预览后生成订单明细；支持扫描文件夹批量处理。 |
| ② 港口 / 数量汇总 | 根据 PO 与订单明细生成港口、数量及特殊港口工作表，汇总价格牌与包装辅料，并可按 PO 输出包装明细 PDF。 |
| ③ 订单汇总 | 按客款号和颜色汇总订单，将不同出厂日期的数量分列展示，生成便于跟进交期的 Excel 工作簿。 |
| ④ 验货报告 | 生成 INLINE 中期报告和 FINAL 尾期报告，同时提供断针记录、BOM 物料分类及验货进度查看。 |
| ⑤ 装箱单生成 | 按款号选择包装方式，配置每箱容量、箱规及重量参数，调整出货日期、成分和包装说明，生成装箱单。 |
| ⑥ 出货表导出 | 扫描装箱单中的目的地，按自定义目的地分组与出货日期导出出货表及点箱数明细。 |
| ⑦ 装箱贴纸 | 核对 SSCC 标签与装箱数据，为标签 PDF 标注箱号，并导出对应明细。 |
| ASOS 标签 / 箱单 | 结合 ASOS PO PDF、工厂装箱资料和 ASN 对照，核对数据并批量生成标签 PDF 与 ASN 箱单 PDF。 |

### 订单与颜色管理

- 自动匹配颜色库和目的港资料，可在界面修正中英文颜色、颜色编号及港口代码。
- 支持单元格编辑、从 Excel 粘贴多行数据，以及将多拼色按件数拆分处理。
- 支持预览 ERP 物料颜色、成品颜色导入清单，调整分类后再提交。
- 对接罗斯 ERP，可查看本次导入结果，并提供删除本次导入的操作。

### 验货与资料交接

- 根据款号与订单数量填充验货报告模板，分别处理 INLINE 与 FINAL 报告。
- 对 `.xlsm` 模板直接修改目标工作表 XML，尽量保留原有宏、图片、公式和格式。
- 按客款号生成 Word 断针记录，从 Tech-Pack `.xls` 中提取并分类汇总 BOM 物料。
- 支持一次触发验货报告、断针记录和 BOM 汇总的生成。
- 通过共享目录中的交接清单与“已完成”文件夹查看批次、款号及待验 / 已完成状态。

## 工具特色

- **一个窗口处理多个环节**：通过标签页切换订单、验货、包装和出货功能，减少在多个独立脚本之间切换。
- **批量处理重复工作**：支持批量导入订单、生成报告与输出标签，适合同一批订单涉及多个款号、颜色和港口的场景。
- **规则可调整**：包装方式、尺码容量、目的地分组和 BOM 分类支持配置或人工修正，便于适配实际业务。
- **先核对再生成**：订单导入提供预览，ERP 导入提供清单核对，标签流程包含装箱数据校验，方便提前发现不匹配的数据。
- **沿用常用文件格式**：输入与输出围绕 CSV、Excel、Word 和 PDF 展开，方便与现有订单资料及模板衔接。
- **本地桌面运行**：常规文件处理在本机完成；ERP 对接需要另行登录并连接对应服务。

## 典型使用流程

1. 准备 PO 合同、订单明细模板、港口表和颜色库，在“订单明细导入”中完成映射并生成明细。
2. 生成港口 / 数量汇总及订单汇总，核对款号、颜色、数量和日期。
3. 按需生成验货报告、断针记录与 BOM 汇总，并通过交接目录跟进验货进度。
4. 配置包装方式与装箱参数，生成装箱单，再按目的地和出货日期导出出货资料。
5. 根据客户要求处理 SSCC 装箱贴纸，或进入 ASOS 专用流程生成标签与 ASN 箱单。

各模块也可按需使用；ASOS 标签 / 箱单与普通包装明细 PDF 是两个独立流程。

## 快速开始

### 使用 Windows EXE

如已取得打包程序，运行 `CO 外贸跟单工具_版本号.exe` 即可打开界面，无需另外安装 Python。首次使用时，在各功能页选择输入文件、输出目录及相应模板。

自定义装箱配置 `packing_config.json` 和 BOM 分类规则 `bom_category_overrides.json` 放在 EXE 同目录；源码运行时则放在 `源代码/` 下。配置文件应使用自己的路径与业务设置。

### 从源码运行

开发环境使用 **Windows + Python 3.11.9**。在项目根目录打开 PowerShell，执行：

```powershell
py -3.11 -m venv .venv
& ./.venv/Scripts/python.exe -m pip install -r requirements-dev.txt
& ./.venv/Scripts/python.exe ./源代码/main_app.py
```

依赖包括 openpyxl、pandas、PyMuPDF、ReportLab、python-docx 和 xlrd 等，具体版本见 `requirements-dev.txt`。

模板和业务资料需自行准备。仅取得公开源码时，验货、断针记录及客户专用流程所需的模板不一定随代码提供；请在对应功能页选择匹配的文件。

## 源码结构

```text
源代码/
├── main_app.py                 # 多功能窗口入口
├── app_version.py              # 版本号与窗口标题
├── po_import_app.py            # 订单明细导入、颜色映射
├── order_erp_client.py         # ERP 颜色库对接
├── erp_credentials.py          # 本机登录信息加密存储
├── generate_tabs.py            # 港口 / 数量汇总、包装明细 PDF
├── order_summary_app.py        # 订单汇总
├── qc_report_app.py            # 验货报告
├── needle_report_app.py        # 断针记录
├── bom_summary_app.py          # BOM 物料分类
├── qc_handoff.py               # 验货资料交接
├── qc_status_app.py            # 验货进度
├── packing_list_app.py         # 装箱单生成
├── shipping_table_app.py       # 出货表界面
├── shipping_table_export.py    # 出货表与点箱数明细导出
├── stamp_cartons_tab.py        # 装箱贴纸界面
├── stamp_cartons.py            # SSCC 校验与箱号标注
├── asos_labels_app.py          # ASOS 标签与 ASN 箱单
└── CO 外贸跟单工具.spec        # Windows 打包配置
requirements-dev.txt           # 依赖清单
```

完整交付目录还可包含 `模板/` 和 `程序/`，用于放置模板、配置与操作说明；它们与公开源码包的内容范围不同。

## 开发与打包

运行现有自动测试：

```powershell
& ./.venv/Scripts/python.exe -m unittest discover -s 源代码 -p 'test_*.py' -v
```

从 `源代码/` 目录构建 Windows EXE：

```powershell
Set-Location ./源代码
& ../.venv/Scripts/python.exe -m PyInstaller --noconfirm 'CO 外贸跟单工具.spec'
```

当前版本的构建结果为 `源代码/dist/CO 外贸跟单工具_1.1.3.exe`。条码字体及其许可通过打包配置一起加入程序；`check_frozen_startup.ps1` 可用于检查 EXE 启动。

版本统一在 `源代码/app_version.py` 的 `APP_VERSION` 中维护，窗口标题、EXE 文件名和 Windows 文件属性使用同一版本。每次发布递增版本号，修订号与次版本号逢 10 进位，例如 `1.1.9 → 1.2.0`、`1.9.9 → 2.0.0`；可通过 `next_version(APP_VERSION)` 计算下一版。

## 账号与数据保护

代码不内置使用者的登录账号、密码或令牌。ERP 登录由使用者自行输入，记住密码功能使用当前 Windows 用户的 DPAPI 加密服务，保存位置为：

```text
%LOCALAPPDATA%/COMerchandiserTool/erp_login.bin
```

该文件属于本机用户数据，不应加入公开仓库或 EXE 发布资源。如果登录框自动填入旧账号，说明读取了此本地缓存；清除该文件后，登录框会恢复为空。

公开发布时，只打包经过检查的源码及必要资源，排除登录缓存、令牌、个人路径配置、生产订单、生成报表、虚拟环境和构建缓存。联调请使用授权的脱敏样本。
