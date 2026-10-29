# dsh-formula-mathtype

DSH 插件：**截图公式 → SimpleTex 识别 → 真正的 MathType 公式 → 插入当前 Word 文档光标处**。

面向论文写作场景：用 `Win+Shift+S` 截下论文里的公式，敲一次 `/formula`，Word 里就多一个可编辑的 MathType 公式对象（`Equation.DSMT4`），不是图片、不是 Word 原生公式。

## 工作流

```
剪贴板截图
   │  helper.py：PIL.ImageGrab 读取
   ▼
SimpleTex 公式识别 API（同步，返回 res.latex + conf）
   │
   ▼
MT6.dll（MathType 自带 SDK 接口）
   MTAPIConnect → MTXFormSetTranslator("LaTeX.tdl") → MTXFormEqn
   │  LaTeX 文本 → 剪贴板上的 MathType OLE 对象
   ▼
Word COM（GetActiveObject("Word.Application") + Selection.Paste）
   │
   ▼
活动文档光标处：可编辑的 MathType 公式
```

实测（本机 MathType 7.11.1 + Word Office16）：热态下 LaTeX→MathType OLE 转换约 **0.04–0.13 秒/条**，Word 粘贴约 0.3–0.5 秒；不弹 MathType 窗口、不运行 Word 宏。完整链路（截图→OCR→插入）冷启动约 7 秒（Python/MathType 引擎启动），之后每条约 **0.4–1.5 秒**（大部分是 OCR 网络时间）。插件加载后会预热 sidecar（常驻 Python + MathType 引擎，约 70–90 MB 内存），换取第一条公式也不慢。

## 组件

| 文件 | 作用 |
| --- | --- |
| `entry.js` | cache-busting 入口，静态声明 `inject = ['tools','commands']` |
| `index.js` | DSH 宿主插件：注册 `/formula` 命令与 `formula_insert` 工具，管理常驻 Python sidecar |
| `helper.py` | Python sidecar：剪贴板取图 / SimpleTex OCR / MT6.dll 转换 / Word COM 粘贴 |
| `cordis.patch.yml` | 插件行与默认配置 |
| `test-harness.mjs` | 不依赖 DSH 的离线自测 |

## 依赖

- Windows（用到剪贴板、Word COM、MathType）
- MathType 7（注册了 `Equation.DSMT4`；`MT6.dll` 随 MathType 安装）
- Microsoft Word 桌面版
- Python 3.10+，且装有 `requests`、`Pillow`、`pywin32`
- 一个 SimpleTex 账户的**用户授权令牌（UAT）**

## 安装

1. 把本项目放到固定目录，例如 `E:\GitHubWorkspace\own\opensource\dsh-formula-mathtype`。
2. 写令牌（一行，前后不要空格）：

   ```
   %LOCALAPPDATA%\dsh-formula-mathtype\simpletex.token
   ```

3. 按本机解释器/令牌路径改 `cordis.patch.yml` 里的 `pythonPath`、`tokenFile`（可选 `model: turbo|standard`）。
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

建议流程：先在 Word 里把光标放到目标段落，再回 DSH 触发；识别结果和置信度会显示在会话里，不满意可直接重截。

## 配置项

| 键 | 默认 | 说明 |
| --- | --- | --- |
| `pythonPath` | `python` | 有 pywin32/Pillow/requests 的解释器绝对路径 |
| `helperPath` | 同目录 `helper.py` | sidecar 路径 |
| `tokenFile` | `%LOCALAPPDATA%\dsh-formula-mathtype\simpletex.token` | 一行一个 UAT |
| `model` | `turbo` | `turbo`=轻量模型（0.04 元/次），`standard`=标准模型（0.05 元/次） |

## 为什么走这条链路

- **不用 Word 宏**：MathType 的 `Toggle TeX` 宏（`MathTypeCommands.UILib.MTCommand_TeXToggle`）在隐藏/自动化 Word 下会挂起并弹 VBA 窗口，无法稳定自动化。
- **不用 OMML/Word 原生公式**：那需要用户再手动跑一次 "Convert Equations" 才能变成 MathType 对象，多一步且容易漏。
- **不用 AddOLEObject + 粘贴面板**：需要激活 MathType 编辑器窗口，慢且会抢焦点。
- **MT6.dll 是 MathType 自带的官方 SDK 接口**：直接生成剪贴板 OLE 对象，静默、快、结果就是真 MathType 对象。

## 令牌与安全

- 令牌只存在本机文件里，插件进程读取后仅用于 `server.simpletex.cn`；不会写日志、不会进入会话记录。
- UAT 是 SimpleTex 标注的“开发用”方式；本地自用足够，若将来分享插件，应改用 APP ID + Secret + MD5 签名。
- 截图会上传到 SimpleTex 服务器；只截公式区域，不要整页带敏感信息的文档。
- 当前按次计费：轻量模型 0.04 元/次、标准 0.05 元/次（前 5000 次/月）；新账户有免费额度。

## 故障排查

| 现象 | 处理 |
| --- | --- |
| `找不到 SimpleTex 令牌文件` | 把 UAT 写进 `tokenFile` 指向的文件 |
| `剪贴板里没有图片` | 先截图（Win+Shift+S / QQ / 微信截图），确认图在剪贴板 |
| `没有找到正在运行的 Word` | 先打开报告文档，把光标放到插入位置 |
| `没有新增 MathType 对象` | 确认光标在正文（不是页眉/图片上），且 MathType 安装完好 |
| `找不到 MathType 的 MT6.dll` | 重装/修复 MathType 7；确认 `HKCR\Equation.DSMT4` 存在 |
| `SimpleTex 识别失败` | 检查令牌、余额、QPS；整页文档图会返回 `[DOCIMG]`，请只截公式 |

## 限制

- 只有 Windows 可用。
- 公式插在 **光标处**，是行内对象；独立公式的居中、编号请沿用你现有的 Word/MathType 排版习惯。
- 依赖桌面 Word（Word 未打开时无法插入；不会替你新建文档）。
- 当前是 A 版（DSH 命令触发）；B 版（全局热键常驻）在路线图中。

## 路线图

- **B**：全局热键（如 `Ctrl+Alt+F`）触发，不切窗口直接插入 Word。
- 可选：MinerU 本地开源版做整篇 PDF 批量导入（不属于即时截图链路）。
- 可选：MathML 后备路径（LaTeX→MathML→MathType），提升个别不支持命令的兼容性。