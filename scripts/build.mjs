import { build } from 'esbuild';
import { mkdir, writeFile } from 'node:fs/promises';
await mkdir(new URL('../dist', import.meta.url), { recursive: true });
const result = await build({ entryPoints: ['src/client.jsx'], bundle: true, write: false, format: 'cjs', platform: 'browser', target: 'es2022', jsx: 'automatic', external: ['react', 'react/*', 'react-dom', 'react-dom/*', '@deepseek-ai/*'], loader: { '.css': 'text' } });
await writeFile('dist/client.js', `// Built from src/client.jsx.\nwindow.__ModuleLoader__.load({id:'dsh-sendfile',factory:(require)=>{var module={exports:{}};var exports=module.exports;\n${result.outputFiles[0].text}\nreturn module.exports;}});\n`);
