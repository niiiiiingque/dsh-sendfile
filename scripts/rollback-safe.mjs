import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
const folder = process.argv[2];
if (!folder || folder.startsWith('--')) throw new Error('用法：node scripts/rollback.mjs <完整备份目录> [--apply]');
const recovery = fileURLToPath(new URL('./recovery.py', import.meta.url));
const args = [recovery, 'restore-profile', folder];
if (process.argv.includes('--apply')) args.push('--apply');
const result = spawnSync('python3', args, { stdio: 'inherit' });
process.exit(result.status ?? 1);
