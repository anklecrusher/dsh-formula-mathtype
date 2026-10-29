/**
 * Cache-busting entry point for the formula-mathtype plugin.
 *
 * Why this file exists
 * --------------------
 * Node caches ES modules by resolved URL for the whole process lifetime. Cordis
 * re-runs a plugin's `apply` on reload, but it cannot force a re-import, so
 * editing `index.js` or `helper.py` would only take effect after a harness
 * restart. This entry imports the implementation through a unique `?t=<time>`
 * URL on each activation instead — the same approach as
 * `@local/dsh-gitbash-tool`. The trade-off is one extra module instance per
 * explicit human activation, not a per-request cost.
 *
 * `inject` must stay a static export: the DI declaration has to be readable
 * without executing the implementation module.
 *
 * @module @local/dsh-formula-mathtype/entry
 */

/** Service dependencies: one tool plus one command. */
export const inject = ['tools', 'commands']

/**
 * Load the implementation fresh and delegate activation to it.
 * @param ctx - plugin context carrying `tools` and `commands`.
 * @param config - plugin config (pythonPath, helperPath, tokenFile, model, hotkey).
 * @returns the implementation's activation result.
 */
export async function apply(ctx, config) {
  const url = new URL(`./index.js?t=${Date.now()}`, import.meta.url)
  const implementation = await import(url.href)
  return implementation.apply(ctx, config)
}