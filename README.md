# CO 外贸跟单工具 · 开发交接

Windows 桌面程序，使用 Python 3.11.9 和 Tkinter。解压后从根目录运行以下 PowerShell 命令：

```powershell
py -3.11 -m venv .venv
& ./.venv/Scripts/python.exe -m pip install -r requirements-dev.txt
& ./.venv/Scripts/python.exe ./源代码/main_app.py
& ./.venv/Scripts/python.exe -m unittest discover -s 源代码 -p 'test_*.py' -v
```

构建 Windows EXE（从源代码目录运行，确保条码字体等资源被打包）：

```powershell
Set-Location ./源代码
& ../.venv/Scripts/python.exe -m PyInstaller --noconfirm 'CO 外贸跟单工具.spec'
```

当前发布版本为 **1.1.2**。版本统一在 `源代码/app_version.py` 的 `APP_VERSION` 中维护，打包文件名、窗口标题和 Windows 文件属性自动使用此版本。后续每次发布必须递增：`1.1.1 → 1.1.2 → … → 1.1.9 → 1.2.0`，修订号和次版本号逢 10 进位（`1.9.9 → 2.0.0`）。可调用该模块的 `next_version(APP_VERSION)` 得到下一版，再更新 `APP_VERSION`。

发布包不得包含罗斯系统账号、密码或登录缓存。登录信息由当前 Windows 用户加密保存在 `%LOCALAPPDATA%/COMerchandiserTool/erp_login.bin`，不属于 EXE 打包资源；本机若自动填入旧账号，是读取了此本地缓存。清除该文件后登录框恢复为空，用户之后自行登录仍可记住其自己的账号密码。发布前检查 EXE 归档没有包含该缓存，也不要将本机用户数据目录加入打包资源。

构建结果在 `源代码/dist/CO 外贸跟单工具_1.1.2.exe`（随版本更新）。需要使用包内装箱配置和 BOM 分类规则时，把生成的 EXE 放到 `程序/` 同目录；源码模式下将这两个 JSON 复制到 `源代码/`。首次打开在界面选择本地输入、输出及模板路径。可在激活虚拟环境后使用 `check_frozen_startup.ps1` 验证 EXE 启动。

- `源代码/`：主程序入口 `main_app.py`、各功能模块、打包配置、字体及许可、自动测试。
- `程序/`：ASOS 操作说明、装箱规格及 BOM 分类规则。
- `模板/`：港口表、验货和断针记录模板。首次使用时可在对应界面选择这些文件。`.xlsm` 为项目自用模板；只在信任来源时启用宏。
- ASOS 官方模板及历史手册另见交接 ZIP；本仓库包含工具的 ASOS 操作说明。

没有附上生产订单、聊天截图、已生成的报表、旧版 EXE、虚拟环境、缓存或本机路径配置。人工联调时请使用授权的脱敏样本。包装明细 PDF 由 `generate_tabs.py` 导出；ASOS 标签及箱单是 `asos_labels_app.py` 的独立流程。
