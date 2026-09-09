import z from '@deepseek-ai/schemastery';
import { DEFAULTS, clampSettings } from './protocol.js';
import { FileStore } from './storage.js';
import { BASE } from './protocol.js';
import { createRouter, ROUTES } from './http.js';
import { toolDefinitions, nativeDefinition } from './tool-definitions.js';
export const name = 'dsh-sendfile';
export const inject = ['webServer', 'workspaceRegistry', 'tools', 'settings'];
export function apply(ctx) {
  const namespace = 'sendfile';
  const section = ctx.settings.register(namespace, z.object({ maxChars: z.number().min(1000).max(200000).default(200000), maxFileMB: z.number().min(1).max(64).default(32), retentionDays: z.number().min(0).max(90).default(7) }), { base: DEFAULTS, validate: clampSettings });
  const store = new FileStore(ctx.workspaceRegistry);
  const handler = createRouter(store, { getSettings: () => clampSettings(section.get()), setSettings: next => ctx.settings.replace(namespace, next) });
  for (const route of ROUTES) ctx.effect(() => ctx.webServer.register({ kind: 'exact', path: `${BASE}/${route}`, handler }), `sendfile ${route}`);
  for (const definition of toolDefinitions(store, () => `http://127.0.0.1:${ctx.webServer.port}`)) ctx.tools.register(nativeDefinition(definition));
  ctx.effect(() => {
    const run = () => store.sweepAll(ctx.workspaceRegistry.list().map(w => w.path), clampSettings(section.get()).retentionDays).catch(() => {});
    const first = setTimeout(run, 15000);
    const timer = setInterval(run, 6 * 3600000);
    return () => { clearTimeout(first); clearInterval(timer); };
  }, 'sendfile cleanup sweep');
}
