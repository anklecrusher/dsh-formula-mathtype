/**
 * dsh-formula-mathtype — screenshot formula -> SimpleTex OCR -> real MathType
 * equation inserted into the active Word document.
 *
 * How the pieces fit
 * ------------------
 * This host plugin owns no Windows automation itself. All of it lives in a
 * persistent Python sidecar (`helper.py serve`), driven over JSON lines:
 *
 *   1. the sidecar reads the formula image from the Windows clipboard,
 *   2. posts it to the SimpleTex formula OCR open API and takes `res.latex`,
 *   3. feeds that LaTeX to MathType's own SDK interface in `MT6.dll`
 *      (`MTAPIConnect` / `MTXFormSetTranslator` / `MTXFormEqn`), which puts a
 *      real MathType OLE object (`Equation.DSMT4`) on the clipboard,
 *   4. pastes it at the caret of the active Word document through Word COM.
 *
 * No MathType window opens, no Word macro runs, and the result is an editable
 * MathType equation object rather than an image or Word-native OMML.
 *
 * Two entry points are registered:
 *   - command `/formula`  — deterministic, no model round-trip:
 *       `/formula`              insert the clipboard screenshot
 *       `/formula \frac{a}{b}`  insert the given LaTeX directly
 *   - tool `formula_insert` — same flow, callable by the model when the user
 *       asks in natural language.
 *
 * The sidecar is started lazily, kept warm across requests (MathType's engine
 * and Word COM startup are paid once), and disposed with the plugin.
 *
 * @module @local/dsh-formula-mathtype
 */

import { spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

/** Directory of this module, used to locate helper.py. */
const MODULE_DIR = dirname(fileURLToPath(import.meta.url))

/** Default Python sidecar next to index.js. */
const DEFAULT_HELPER_PATH = join(MODULE_DIR, 'helper.py')

/** Default one-line SimpleTex token file, outside the repository. */
const DEFAULT_TOKEN_FILE = process.env.LOCALAPPDATA
  ? join(process.env.LOCALAPPDATA, 'dsh-formula-mathtype', 'simpletex.token')
  : ''

/**
 * Probe order for a Python interpreter when the config does not name one.
 * GUI processes often have no developer PATH, so probing the common install
 * locations is more reliable than shelling out to a bare `python`.
 * @returns {string[]}
 */
function pythonCandidates() {
  const out = []
  const push = (value) => {
    if (typeof value !== 'string') return
    const trimmed = value.trim()
    if (trimmed.length > 0 && !out.includes(trimmed)) out.push(trimmed)
  }
  push(process.env.DSH_FORMULA_PYTHON)
  push(process.env.PYTHON)
  const localAppData = process.env.LOCALAPPDATA
  if (localAppData) {
    for (const version of ['Python313', 'Python312', 'Python311', 'Python310']) {
      push(join(localAppData, 'Programs', 'Python', version, 'python.exe'))
    }
  }
  for (const version of ['313', '312', '311', '310']) {
    push(`C:/Python${version}/python.exe`)
    if (process.env.ProgramFiles) push(join(process.env.ProgramFiles, `Python${version}`, 'python.exe'))
  }
  // This machine keeps its real Python in a non-standard directory; keep it as
  // a last-resort probe so the plugin works without extra config.
  push('D:/Program Files/python.exe')
  return out
}

/**
 * Resolve the Python interpreter: explicit config wins, then known locations,
 * then whatever `python` means on PATH.
 * @param {Record<string, any>} [config]
 * @returns {string}
 */
function resolvePythonPath(config) {
  if (typeof config?.pythonPath === 'string' && config.pythonPath.trim().length > 0) {
    return config.pythonPath.trim()
  }
  for (const candidate of pythonCandidates()) {
    if (existsSync(candidate)) return candidate
  }
  return 'python'
}

/** One formula should never take longer than this; OCR + Word COM included. */
const REQUEST_TIMEOUT_MS = 180_000

/** Timeout for the non-blocking MathType engine warm-up at plugin load. */
const WARMUP_TIMEOUT_MS = 60_000

/**
 * Persistent JSON-lines client for helper.py.
 *
 * The child is started on the first request and kept alive afterwards. A
 * request/response pair is matched by an incrementing `id`; messages without a
 * known id (diagnostics accidentally printed to stdout) are ignored.
 */
class PythonSidecar {
  /**
   * @param {{pythonPath: string, helperPath: string, tokenFile: string, model: string}} options
   */
  constructor(options) {
    this.options = options
    this.child = null
    this.buffer = ''
    this.stderrTail = ''
    this.pending = new Map()
    this.nextId = 1
    this.disposed = false
    // Last-resort cleanup: if the host process exits while the sidecar is still
    // alive (app quit, crash), take the child down with it. Without this the
    // Python process and MathType's engine can outlive the app.
    this.exitHook = () => this.killChild()
    process.once('exit', this.exitHook)
  }

  /** True when the child process is alive and usable. */
  get running() {
    return this.child !== null && this.child.exitCode === null
  }

  /** Spawn the sidecar if it is not running yet. */
  start() {
    if (this.disposed) throw new Error('公式 sidecar 已关闭')
    if (this.running) return

    const args = [
      this.options.helperPath,
      'serve',
      '--token-file',
      this.options.tokenFile,
      '--model',
      this.options.model,
    ]
    const child = spawn(this.options.pythonPath, args, {
      cwd: dirname(this.options.helperPath),
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true,
      env: {
        ...process.env,
        PYTHONIOENCODING: 'utf-8',
        PYTHONUTF8: '1',
      },
    })

    this.child = child
    this.buffer = ''
    this.stderrTail = ''

    child.stdout.setEncoding('utf8')
    child.stdout.on('data', (chunk) => this.onStdout(chunk))
    child.stderr.setEncoding('utf8')
    child.stderr.on('data', (chunk) => {
      this.stderrTail = (this.stderrTail + chunk).slice(-4000)
    })
    child.on('error', (error) => {
      this.child = null
      const hint =
        error?.code === 'ENOENT'
          ? `（找不到 Python 解释器 "${this.options.pythonPath}"，请在插件配置 pythonPath 里填 python.exe 的绝对路径）`
          : ''
      this.failAll(new Error(`公式 sidecar 启动失败：${error.message}${hint}`))
    })
    child.on('exit', (code, signal) => {
      this.child = null
      this.failAll(
        new Error(`公式 sidecar 进程退出 (code=${code}, signal=${signal})：${this.stderrTail.slice(-1500)}`),
      )
    })
  }

  /** Feed stdout bytes through the line protocol. @param {string} chunk */
  onStdout(chunk) {
    this.buffer += chunk
    let newline = this.buffer.indexOf('\n')
    while (newline >= 0) {
      const line = this.buffer.slice(0, newline).trim()
      this.buffer = this.buffer.slice(newline + 1)
      newline = this.buffer.indexOf('\n')
      if (line.length === 0) continue
      let message
      try {
        message = JSON.parse(line)
      } catch {
        continue
      }
      if (typeof message?.id !== 'number') continue
      const pending = this.pending.get(message.id)
      if (pending === undefined) continue
      this.pending.delete(message.id)
      clearTimeout(pending.timer)
      pending.resolve(message)
    }
  }

  /** Reject every in-flight request; used when the child dies or is disposed. @param {Error} error */
  failAll(error) {
    for (const pending of this.pending.values()) {
      clearTimeout(pending.timer)
      pending.reject(error)
    }
    this.pending.clear()
  }

  /**
   * Send one request and resolve with the parsed response envelope.
   * A response with `ok: false` still resolves; the caller decides how to report it.
   * @param {Record<string, unknown>} payload
   * @param {number} [timeoutMs]
   * @returns {Promise<Record<string, any>>}
   */
  request(payload, timeoutMs = REQUEST_TIMEOUT_MS) {
    if (this.disposed) return Promise.reject(new Error('公式 sidecar 已关闭'))
    try {
      this.start()
    } catch (error) {
      return Promise.reject(error)
    }
    const child = this.child
    if (child === null) return Promise.reject(new Error('公式 sidecar 未运行'))

    const id = this.nextId++
    const line = `${JSON.stringify({ id, ...payload })}\n`
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id)
        // A hung sidecar cannot be reused safely: reset it so the next call starts clean.
        this.reset()
        reject(new Error(`公式 sidecar 超时（${timeoutMs} ms）`))
      }, timeoutMs)
      this.pending.set(id, { resolve, reject, timer })
      child.stdin.write(line, 'utf8')
    })
  }

  /** Kill the child process and forget it. */
  reset() {
    const child = this.child
    this.child = null
    this.failAll(new Error('公式 sidecar 已重置'))
    if (child !== null) {
      try {
        child.stdin.write(`${JSON.stringify({ id: 0, action: 'quit' })}\n`)
      } catch {
        /* already closed */
      }
      try {
        child.kill()
      } catch {
        /* already gone */
      }
    }
  }

  /** Kill the child immediately (no graceful quit). */
  killChild() {
    const child = this.child
    this.child = null
    if (child === null) return
    try {
      child.kill()
    } catch {
      /* already gone */
    }
  }

  /** Graceful shutdown, called when the plugin is disposed. */
  dispose() {
    if (this.disposed) return
    this.disposed = true
    const child = this.child
    this.child = null
    this.failAll(new Error('公式 sidecar 已关闭'))
    if (child === null) return
    try {
      child.stdin.write(`${JSON.stringify({ id: 0, action: 'quit' })}\n`)
    } catch {
      /* already closed */
    }
    const timer = setTimeout(() => {
      try {
        child.kill()
      } catch {
        /* already gone */
      }
    }, 800)
    if (typeof timer.unref === 'function') timer.unref()
  }
}

/**
 * Pick the effective mode for a call.
 * @param {Record<string, any>} args
 * @returns {'clipboard'|'latex'|'image'}
 */
function resolveMode(args) {
  if (typeof args?.mode === 'string' && args.mode.trim().length > 0) {
    const mode = args.mode.trim().toLowerCase()
    if (mode === 'clipboard' || mode === 'latex' || mode === 'image') return mode
    throw new Error(`未知 mode "${args.mode}"：只支持 clipboard / latex / image`)
  }
  if (typeof args?.latex === 'string' && args.latex.trim().length > 0) return 'latex'
  if (typeof args?.imagePath === 'string' && args.imagePath.trim().length > 0) return 'image'
  return 'clipboard'
}

/**
 * Validate arguments before touching the sidecar, so obvious mistakes fail fast.
 * @param {'clipboard'|'latex'|'image'} mode
 * @param {Record<string, any>} args
 */
function validateArgs(mode, args) {
  if (mode === 'latex' && (typeof args?.latex !== 'string' || args.latex.trim().length === 0)) {
    throw new Error('mode=latex 时必须在 latex 参数里给出公式文本')
  }
  if (mode === 'image' && (typeof args?.imagePath !== 'string' || args.imagePath.trim().length === 0)) {
    throw new Error('mode=image 时必须在 imagePath 参数里给出图片绝对路径')
  }
}

/**
 * Format a normalized sidecar result for the model and the transcript.
 * @param {Record<string, any>} result
 * @returns {string}
 */
function describeResult(result) {
  const conf = typeof result.conf === 'number' ? result.conf.toFixed(3) : 'n/a'
  const elapsed = typeof result.elapsedMs === 'number' ? result.elapsedMs : 0
  return `已插入 MathType 公式（${elapsed} ms，置信度 ${conf}）：\n${result.latex ?? ''}`
}

/**
 * Build the model-facing tool definition.
 * @param {PythonSidecar} sidecar
 */
function createTool(sidecar) {
  return {
    name: 'formula_insert',
    description:
      '把公式截图识别成 LaTeX，并作为真正的 MathType 公式插入当前打开的 Word 文档光标处。'
      + '默认从 Windows 剪贴板读取截图（用户先用 Win+Shift+S 或截图工具截取公式）。'
      + '也可以传 latex 直接插入指定公式，或传 imagePath 指定图片文件。'
      + '只用于数学公式，不要用它插入普通文字。返回识别出的 LaTeX、置信度和耗时。',
    parameters: {
      type: 'object',
      properties: {
        mode: {
          type: 'string',
          description: 'clipboard（默认，读剪贴板截图）/ latex（直接给公式文本）/ image（给 imagePath 图片路径）',
        },
        latex: {
          type: 'string',
          description: 'mode=latex 时要插入的 LaTeX 文本',
        },
        imagePath: {
          type: 'string',
          description: 'mode=image 时公式图片的绝对路径',
        },
      },
      required: [],
      additionalProperties: false,
    },
    output: {
      schema: {
        type: 'object',
        properties: {
          ok: { type: 'boolean', description: '是否成功插入' },
          latex: { type: 'string', description: '识别（或传入）的 LaTeX' },
          conf: {
            oneOf: [{ type: 'number' }, { type: 'null' }],
            description: 'SimpleTex 识别置信度；mode=latex 时为 null',
          },
          inserted: { type: 'integer', description: '本次插入新增的 MathType OLE 对象数量' },
          elapsedMs: { type: 'integer', description: '总耗时（毫秒）' },
          message: { type: 'string', description: '结果说明' },
        },
        required: ['ok', 'latex', 'conf', 'inserted', 'elapsedMs', 'message'],
        additionalProperties: false,
      },
      render(_args, value) {
        return [{ type: 'text', text: describeResult(value) }]
      },
    },
    async execute(args) {
      const mode = resolveMode(args)
      validateArgs(mode, args)
      const result = await sidecar.request({
        action: 'run',
        mode,
        latex: args.latex,
        imagePath: args.imagePath,
      })
      if (result.ok !== true) {
        throw new Error(typeof result.error === 'string' ? result.error : '公式插入失败')
      }
      return {
        ok: true,
        latex: typeof result.latex === 'string' ? result.latex : '',
        conf: typeof result.conf === 'number' ? result.conf : null,
        inserted: typeof result.inserted === 'number' ? result.inserted : 1,
        elapsedMs: typeof result.elapsedMs === 'number' ? result.elapsedMs : 0,
        message: typeof result.message === 'string' ? result.message : '已插入 MathType 公式',
      }
    },
  }
}

/**
 * Build the `/formula` command definition (deterministic, no model round-trip).
 * @param {PythonSidecar} sidecar
 */
function createCommand(sidecar) {
  return {
    name: 'formula',
    description: '识别剪贴板里的公式截图并以 MathType 公式插入当前 Word 文档；也可 /formula <LaTeX> 直接插入',
    input: {
      hint: '[可选的 LaTeX 文本；省略则读取剪贴板截图]',
    },
    async handler(invocation) {
      const raw = String(invocation?.rawInput ?? '').trim()
      const mode = raw.length > 0 ? 'latex' : 'clipboard'
      try {
        const result = await sidecar.request({
          action: 'run',
          mode,
          latex: mode === 'latex' ? raw : undefined,
        })
        if (result.ok !== true) {
          throw new Error(typeof result.error === 'string' ? result.error : '公式插入失败')
        }
        return { kind: 'success', text: describeResult(result) }
      } catch (error) {
        return { kind: 'error', text: `公式插入失败：${error instanceof Error ? error.message : String(error)}` }
      }
    },
  }
}

/**
 * Merge plugin config with defaults.
 * @param {Record<string, any>} [config]
 */
function resolveOptions(config) {
  const pythonPath = resolvePythonPath(config)
  const helperPath =
    typeof config?.helperPath === 'string' && config.helperPath.trim().length > 0
      ? config.helperPath.trim()
      : DEFAULT_HELPER_PATH
  const tokenFile =
    typeof config?.tokenFile === 'string' && config.tokenFile.trim().length > 0
      ? config.tokenFile.trim()
      : DEFAULT_TOKEN_FILE
  const model = config?.model === 'standard' ? 'standard' : 'turbo'
  return { pythonPath, helperPath, tokenFile, model }
}

/**
 * Plugin entry point.
 * @param ctx - plugin context; requires `tools` and `commands`.
 * @param config - plugin config: pythonPath, helperPath, tokenFile, model.
 */
export function apply(ctx, config) {
  const options = resolveOptions(config)
  const sidecar = new PythonSidecar(options)

  if (typeof ctx.effect === 'function') {
    ctx.effect(() => () => sidecar.dispose(), 'dsh-formula-mathtype: python sidecar')
  }

  ctx.tools.register(createTool(sidecar))
  ctx.commands.register(createCommand(sidecar))

  // Warm the sidecar (Python + MathType engine) without blocking activation.
  // Word is intentionally not touched: it only exists when the user has it open.
  sidecar.request({ action: 'warmup' }, WARMUP_TIMEOUT_MS).catch(() => {})
}

/** Package identity, for diagnostics. */
export const name = 'dsh-formula-mathtype'