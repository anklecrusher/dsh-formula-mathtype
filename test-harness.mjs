/**
 * Standalone harness: exercise the plugin without the DSH runtime.
 *
 * Mocks ctx.tools.register / ctx.commands.register / ctx.effect, verifies what
 * was registered, checks argument validation, drives sidecar requests through the
 * command handler (in DSH_FORMULA_TEST_MODE they come back as deterministic
 * errors, so no OCR/insertion can happen), runs the offline Python selftest, and
 * finally checks the `/formula hotkey status` subcommand of the B-mode companion.
 *
 * Run with: node test-harness.mjs
 */
import { spawnSync } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { apply } from './index.js'

const HERE = dirname(fileURLToPath(import.meta.url))

// Never let a test run touch the real clipboard, spend an OCR call, or insert a
// formula into a Word/PowerPoint document the user has open. helper.py returns a
// deterministic "test mode" error instead.
process.env.DSH_FORMULA_TEST_MODE = '1'

let registeredTool = null
let registeredCommand = null
let dispose = null

const ctx = {
  tools: {
    register(definition) {
      registeredTool = definition
      return () => {}
    },
  },
  commands: {
    register(definition) {
      registeredCommand = definition
      return () => {}
    },
  },
  effect(callback) {
    dispose = callback()
    return dispose
  },
}

apply(ctx, {})

let failures = 0
const check = (label, condition, detail) => {
  if (condition) {
    console.log(`  PASS  ${label}`)
  } else {
    failures += 1
    console.log(`  FAIL  ${label} :: ${detail}`)
  }
}

console.log('registered tool    :', registeredTool?.name)
console.log('registered command :', registeredCommand?.name)
console.log('---')

check('tool registered', registeredTool !== null, 'no tool')
check('tool name', registeredTool?.name === 'formula_insert', String(registeredTool?.name))
check('tool params', JSON.stringify(Object.keys(registeredTool?.parameters?.properties ?? {})) === '["mode","latex","imagePath","target"]', JSON.stringify(registeredTool?.parameters))
check('tool target enum', JSON.stringify(registeredTool?.parameters?.properties?.target?.enum) === '["word","ppt"]', JSON.stringify(registeredTool?.parameters?.properties?.target))
check('command registered', registeredCommand !== null, 'no command')
check('command name', registeredCommand?.name === 'formula', String(registeredCommand?.name))
check('command has handler', typeof registeredCommand?.handler === 'function', typeof registeredCommand?.handler)

const expectReject = async (label, promise) => {
  try {
    await promise
    check(label, false, 'resolved unexpectedly')
  } catch (error) {
    check(label, error instanceof Error && error.message.length > 0, String(error))
  }
}

await expectReject('mode=latex without latex rejected', registeredTool.execute({ mode: 'latex' }))
await expectReject('mode=image without imagePath rejected', registeredTool.execute({ mode: 'image' }))
await expectReject('unknown mode rejected', registeredTool.execute({ mode: 'clipboard2' }))
await expectReject('unknown target rejected', registeredTool.execute({ mode: 'latex', latex: 'x', target: 'excel' }))

// Offline Python selftest: normalize_latex / brace detection / hotkey parsing /
// test-mode guard. Kept in helper.py so it can import the real functions.
const python = [process.env.DSH_FORMULA_TEST_PYTHON, process.env.DSH_FORMULA_PYTHON, 'D:/Program Files/python.exe', 'python3', 'python']
  .filter((candidate) => typeof candidate === 'string' && candidate.length > 0)
  .find((candidate) => spawnSync(candidate, ['-c', 'import sys'], { windowsHide: true }).error === undefined)
if (python) {
  const selftest = spawnSync(python, [join(HERE, 'helper.py'), 'selftest'], { encoding: 'utf8', windowsHide: true })
  check('python helper selftest', selftest.status === 0, String(selftest.stdout ?? selftest.stderr ?? selftest.error).slice(-400))
  if (selftest.status !== 0) console.log(String(selftest.stdout ?? '').split('\n').slice(-12).join('\n'))
} else {
  check('python helper selftest', false, 'no python interpreter found')
}

// Sidecar round trip: in test mode helper.py must refuse before reading the
// clipboard or inserting anything, and the command must report it cleanly.
const commandResult = await registeredCommand.handler({ rawInput: '' })
check('command test-mode error returned', commandResult?.kind === 'error', JSON.stringify(commandResult))
check('command test-mode error text', String(commandResult?.text ?? '').includes('测试模式'), JSON.stringify(commandResult))
console.log('  sidecar error text :', String(commandResult?.text ?? '').split('\n')[0])

// PowerPoint routing: /formula ppt must reach the sidecar with target=ppt.
const pptResult = await registeredCommand.handler({ rawInput: 'ppt' })
check(
  'ppt command reaches sidecar',
  pptResult?.kind === 'error' && String(pptResult?.text ?? '').includes('测试模式'),
  JSON.stringify(pptResult),
)
console.log('  ppt command text  :', String(pptResult?.text ?? '').split('\n')[0])

// B mode: the hotkey companion starts after the sidecar warm-up. Give it a
// moment, then verify the status subcommand wired into /formula.
await new Promise((resolve) => setTimeout(resolve, 2500))
const hotkeyStatus = await registeredCommand.handler({ rawInput: 'hotkey status' })
const hotkeyText = String(hotkeyStatus?.text ?? '')
const externalHelper = hotkeyText.includes('另一个热键助手已在运行')
check(
  'hotkey status command',
  hotkeyStatus?.kind === 'success' &&
    hotkeyText.includes('热键') &&
    hotkeyText.includes('Ctrl+Alt+F') &&
    (hotkeyText.includes('运行中') || externalHelper),
  JSON.stringify(hotkeyStatus),
)
if (externalHelper) console.log('  (skipped: a DSH-managed hotkey helper is already running)')
console.log('  hotkey status text :', String(hotkeyStatus?.text ?? '').split('\n').join(' / '))
const hotkeyOff = await registeredCommand.handler({ rawInput: 'hotkey off' })
check('hotkey off command', hotkeyOff?.kind === 'success', JSON.stringify(hotkeyOff))

if (typeof dispose === 'function') dispose()
await new Promise((resolve) => setTimeout(resolve, 500))

console.log('---')
if (failures > 0) {
  console.error(`FAIL: ${failures} check(s) failed`)
  process.exit(1)
}
console.log('OK: harness passed')