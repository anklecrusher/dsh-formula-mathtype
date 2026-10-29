/**
 * Standalone harness: exercise the plugin without the DSH runtime.
 *
 * Mocks ctx.tools.register / ctx.commands.register / ctx.effect, verifies what
 * was registered, checks argument validation, and then drives one real sidecar
 * request through the command handler (expected to fail cleanly here because no
 * Word document is open in the harness environment — the point is that the
 * JSON-lines round trip and the error path work, not the insertion itself).
 *
 * Run with: node test-harness.mjs
 */
import { apply } from './index.js'

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
check('tool params', JSON.stringify(Object.keys(registeredTool?.parameters?.properties ?? {})) === '["mode","latex","imagePath"]', JSON.stringify(registeredTool?.parameters))
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

// Real sidecar round trip: with no clipboard image (or no Word) this must come
// back as a clean {kind:'error'} result, not a crash.
const commandResult = await registeredCommand.handler({ rawInput: '' })
check('command returns structured result', commandResult?.kind === 'success' || commandResult?.kind === 'error', JSON.stringify(commandResult))
check('command error path is reported', commandResult?.kind === 'error', JSON.stringify(commandResult))
if (commandResult?.kind === 'error') {
  console.log('  sidecar error text :', String(commandResult.text).split('\n')[0])
}

if (typeof dispose === 'function') dispose()
await new Promise((resolve) => setTimeout(resolve, 300))

console.log('---')
if (failures > 0) {
  console.error(`FAIL: ${failures} check(s) failed`)
  process.exit(1)
}
console.log('OK: harness passed')