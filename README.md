# dsh-formula-mathtype

DSH 插件：**截图公式 → SimpleTex 识别 → 真正的 MathType 公式 → 插入当前 Word 文档光标处**。

面向论文写作场景：用 `Win+Shift+S` 截下论文里的公式，敲一次 `/formula`，Word 里就多一个可编辑的 MathType 公式对象（`Equation.DSMT4`），不是图片、不是 Word 原生公式。

## 工作流

```
剪贴板截图
   │  helper.py：PIL.ImageGrab 读取
   ▼
SimpleTex 公式识别 API（standard 模型，同步返回 res.latex + conf）
   │
   ▼
LaTeX 归一化（helper.py）
   aligned/align/split → array；保留 cases/array/matrix；
   从原图检测大括号 → \left\{\begin{array}{l}...\end{array}\right.
   │
   ▼
injector/MathTypeInjector.exe（Word COM + MathType OLE）
   InlineShapes.AddOLEObject("Equation.DSMT4") → OLEFormat.DoVerb
   → IDataObject.SetData("TeX Input Language", LaTeX) → IOleObject.Close
   │
   ▼
活动文档光标处：可编辑的 MathType 公式
```

不弹 MathType 窗口、不运行 Word 宏、不依赖 Office PIA。注入器是 C# 源码，首次使用由 `helper.py` 调 Windows 自带的 `.NET Framework` 编译器（`csc.exe`）自动构建。

## 实测性能（本机 MathType 7.11.1 + Word Office16）

| 场景 | 耗时 |
| --- | --- |
| 单行简单公式（截图→OCR→插入） | 约 2.5 秒 |
| 6 行大括号方程组（standard 模型） | OCR 约 10 秒 + 插入约 2.7 秒 |
| B 模式按热键到插入（简单公式） | 约 1.2–2.5 秒 |

耗时主要在 SimpleTex 标准模型的网络识别上；插入本身（OLE 对象创建 + TeX 解析 + 保存）对简单公式约 0.4–0.7 秒、对 6 行方程组约 2.7 秒。

## 复杂公式支持（v0.2.0）

- **多行方程组**：SimpleTex 常把大括号系统识别成 `\begin{aligned}...` 并丢掉大括号。插件会把 `aligned` 归一化成 MathType 支持的 `array`，并根据原图左侧是否存在高瘦的大括号，决定是否加回 `\left\{...\right.`。
- **MathType 不支持的写法**：`aligned`/`align`/`split`/`gathered` 会被 MathType 渲染成红色错误符号，已在归一化阶段转换；`array`/`matrix`/`cases`/`\left\{...\right.` 经实测可用。
- **识别模型**：默认 `standard`（0.05 元/次）。`turbo`（0.04 元/次）更快，但实测对多行方程组会出现重复行、括号不平衡，复杂公式不要用。
- OCR 仍可能有个别下标识别错误（例如把 `\xi_{\psi i}` 认成 `\xi_{yi}`）；插入后可直接在 MathType 里改，不影响对象可用性。

## 组件

| 文件 | 作用 |
| --- | --- |
| `entry.js` | cache-busting 入口，静态声明 `inject = ['tools','commands']` |
| `index.js` | DSH 宿主插件：注册 `/formula` 命令与 `formula_insert` 工具，管理常驻 Python sidecar |
| `helper.py` | Python sidecar：剪贴板取图 / SimpleTex OCR / LaTeX 归一化 / 大括号检测 / 调注入器 |
| `injector/MathTypeInjector.cs` | Word COM + MathType OLE 注入器源码（首次使用自动编译） |
| `hotkey_runner.py` | B 模式常驻助手：全局热键 + 原生 Win32 toast + 复用插入逻辑 |
| `cordis.patch.yml` | 插件行与默认配置 |
| `test-harness.mjs` | 不依赖 DSH 的离线自测 |

## 依赖

- Windows（剪贴板、Word COM、MathType OLE）
- MathType 7（注册了 `Equation.DSMT4` 的 OLE 服务器即可，不需要 SDK）
- Microsoft Word 桌面版
- .NET Framework 4.x（Windows 自带 `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`，用于首次构建注入器）
- Python 3.10+，装有 `requests`、`Pillow`、`numpy`（`pywin32` 只在失败回退时用于把 LaTeX 复制到剪贴板，可选）
- B 模式的热键提示用原生 Win32 分层窗口绘制，**不需要 Tcl/Tk**（本机 Python 的 Tcl 数据不完整，Tkinter 不可用也能工作）

## 安装

1. 把本项目放到固定目录，例如 `E:\GitHubWorkspace\own\opensource\dsh-formula-mathtype`。
2. 写令牌（一行，前后不要空格）：

   ```
   %LOCALAPPDATA%\dsh-formula-mathtype\simpletex.token
   ```

3. 按本机解释器/令牌路径改 `cordis.patch.yml` 里的 `pythonPath`、`tokenFile`（可选 `model: standard|turbo`）。
4. 加入 desktop profile（desktop profile 只能由桌面端自带 CLI 管理）：

   ```powershell
   & "D:\Users\ankle crusher\AppData\Local\Programs\DeepSeek Harness\resources\runtime\cli\bin\dsh.cmd" `
       plugin --profile desktop add "@local/dsh-formula-mathtype@link:E:/GitHubWorkspace/own/opensource/dsh-formula-mathtype"
   ```

   > 注意：这个版本的 `dsh plugin add` 只是把参数转发给 pnpm，**不会自动写 `dsh.profile.bundles`**。安装成功后还要把 `"@local/dsh-formula-mathtype"` 加进 profile `package.json` 的 `dsh.profile.bundles`（放在 `@local/dsh-latex-tool` 后面即可）。
   >
   > 若 pnpm 11 报 `ERR_PNPM_IGNORED_BUILDS`（本机是 `protobufjs`），先在 profile 目录批准一次构建：
   > ```powershell
   > & "D:\Users\ankle crusher\AppData\Local\Programs\DeepSeek Harness\DeepSeek Harness.exe" --expose-internals `
   >   "D:\Users\ankle crusher\AppData\Local\Programs\DeepSeek Harness\resources\runtime\primary-runtime\dependencies\pnpm\bin\pnpm.mjs" `
   >   approve-builds protobufjs
   > ```
   > 它会在 `profile/pnpm-workspace.yaml` 里写入 `allowBuilds: { protobufjs: true }`；不修的话之后任何插件安装都会被这个严格构建检查挡住。

5. 重启 DSH 桌面端（bundle 行在启动时读取），之后 `/` 命令列表里会出现 `formula`。

## 用法

- 截图后（图在剪贴板里），在 DSH 里输入 `/formula` 回车 → 插入当前 Word 文档光标处。
- 也可以直接给 LaTeX：`/formula \frac{a}{b}=\sqrt{x^{2}+y^{2}}`。
- 自然语言也可以：让模型调用 `formula_insert` 工具（参数 `mode: clipboard|latex|image`）。

建议流程：先在 Word 里把光标放到目标段落，再回 DSH 触发；识别结果、置信度和耗时会在会话里显示。多行方程组会自动还原大括号。

## B 模式：全局热键（v0.3.0）

插件加载后会常驻一个热键小助手（`helper.py hotkey`），注册系统级热键（默认 `Ctrl+Alt+F`）。使用流程：

1. `Win+Shift+S` 截公式（图进剪贴板）；
2. **保持在 Word 里**，直接按 `Ctrl+Alt+F`；
3. 右上角弹出小提示「正在识别公式…」，完成后变成「✓ 已插入 MathType 公式」并显示 LaTeX 预览，2.6 秒后自动消失；
4. 失败时显示红色错误提示（8 秒），详情写入 `%LOCALAPPDATA%\dsh-formula-mathtype\hotkey.log`。

实测：简单公式从按键到插入约 **1.2–2.5 秒**，6 行方程组约 10 秒（主要是标准模型识别时间）。提示窗口是无激活（`WS_EX_NOACTIVATE`）的原生 Win32 分层窗口，不会抢 Word 焦点；本机 Python 的 Tcl/Tk 数据不完整，所以没有用 Tkinter。

热键用低级键盘钩子（`WH_KEYBOARD_LL`）实现：按下 `Ctrl+Alt+F` 时组合键会被吞掉，**不会触发 Word 自带的「插入脚注」**（Word 里 `Ctrl+Alt+F` 正是插入脚注快捷键，已实测确认）。因为组合键不含单独的 Alt，不会触发 Word 的键位提示。

副作用：插件运行期间，Word 的 `Ctrl+Alt+F` 插入脚注快捷键会被本插件占用；需要插脚注时用「引用 → 插入脚注」功能区按钮。

启用/关闭/查看状态：

```
/formula hotkey            # 查看运行状态、热键、最近错误
/formula hotkey on         # 启动热键助手
/formula hotkey off        # 停止热键助手
/formula hotkey restart    # 重启（改配置后可用）
```

配置 `hotkey: ''` 可以完全不启动热键助手；热键被其它软件占用时换一个（如 `Ctrl+Alt+M`）。

## 配置项

| 键 | 默认 | 说明 |
| --- | --- | --- |
| `pythonPath` | 自动探测 | 有 requests/Pillow/numpy 的解释器绝对路径 |
| `helperPath` | 同目录 `helper.py` | sidecar 路径 |
| `tokenFile` | `%LOCALAPPDATA%\dsh-formula-mathtype\simpletex.token` | 一行一个 UAT |
| `model` | `standard` | `standard`=标准模型（0.05 元/次，复杂公式明显更准），`turbo`=轻量模型（0.04 元/次，复杂方程组不稳） |
| `hotkey` | `Ctrl+Alt+F` | B 模式全局热键；写成空字符串 `''` 关闭热键常驻 |

## 为什么走这条链路

- **不用 MathType SDK 的 `MTXFormEqn` 反向写公式**：它是导出接口（MathType 对象→文本）。把文本塞进去反向使用，Word 里嵌入的对象内容是字面 LaTeX，渲染成源码或黑块。这是 v0.1.0 的病根。
- **不用 Word 宏**：MathType 的 `Toggle TeX` 宏在自动化 Word 下会挂起并弹 VBA 窗口。
- **不用 OMML/Word 原生公式**：还需要用户手动跑一次 "Convert Equations" 才能变 MathType 对象。
- **现在的方式**：Word 自己创建 `Equation.DSMT4` OLE 对象，再通过该对象标准的 `IDataObject` 接口喂 `TeX Input Language`（MathType 的 TeX 解析器），最后 `IOleObject.Close` 保存。静默、无焦点抢占，结果是真 MathType 对象。

## 令牌与安全

- 令牌只存在本机文件里，插件进程读取后仅用于 `server.simpletex.cn`；不会写日志、不会进入会话记录。
- UAT 是 SimpleTex 标注的“开发用”方式；本地自用足够，若将来分享插件，应改用 APP ID + Secret + MD5 签名。
- 截图会上传到 SimpleTex 服务器；只截公式区域，不要整页带敏感信息的文档。
- 当前按次计费：标准模型 0.05 元/次、轻量模型 0.04 元/次（前 5000 次/月）；新账户有免费额度。

## 故障排查

| 现象 | 处理 |
| --- | --- |
| Word 里出现黑块/源码 | 是 v0.1.0 的旧插入路径产物：升级后重启 DSH，删掉旧对象重新截 |
| `找不到 SimpleTex 令牌文件` | 把 UAT 写进 `tokenFile` 指向的文件 |
| `剪贴板里没有图片` | 先截图（Win+Shift+S / QQ / 微信截图），确认图在剪贴板 |
| `没有找到正在运行的 Word` | 先打开报告文档，把光标放到插入位置 |
| `Word 里没有打开的文档` | 先打开/新建文档 |
| `构建 MathType 注入器失败` | 确认 `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe` 存在；删除 `injector/MathTypeInjector.exe` 后重试 |
| `SimpleTex 识别失败` | 检查令牌、余额、QPS；整页文档图会返回 `[DOCIMG]`，请只截公式区域 |
| 公式渲染出红色错误符号 | 说明还有 MathType TeX 不支持的写法：把该 LaTeX 发给我，加入归一化规则 |
| `/formula hotkey` 显示未运行 | 看状态里的最近错误；常见是键盘钩子安装失败（安全软件/输入法抢占），换一个 `hotkey` 后重启 DSH |
| 提示“另一个热键助手已在运行” | 上次 DSH 异常退出留下的进程：任务管理器结束对应 `python.exe`（命令行含 `helper.py hotkey`）后 `/formula hotkey restart` |
| 按热键没反应 | 先在 Word 里确认光标位置；再 `/formula hotkey status` 看助手是否在运行；详细过程见 `hotkey.log` |
| 按 `Ctrl+Alt+F` 后 Word 多出一个脚注 | 说明键盘钩子没生效（助手未运行或被安全软件拦截），`/formula hotkey status` 查看 |

## 限制

- 只有 Windows 可用。
- 公式插在**光标处**，是行内对象；独立公式的居中、编号请沿用你现有的 Word/MathType 排版习惯。
- 依赖桌面 Word（Word 未打开时无法插入；不会替你新建文档）。
- 多行方程组识别约 10 秒（标准模型网络耗时）。
- A 版（`/formula` 命令）与 B 版（全局热键）都已可用；B 版会常驻一个 Python 热键进程（约 30–40 MB 内存），配置 `hotkey: ''` 可关闭。

## 路线图

- 公式居中/编号快捷处理（按现有 Word 样式）。
- 可选：MinerU 本地开源版做整篇 PDF 批量导入（不属于即时截图链路）。