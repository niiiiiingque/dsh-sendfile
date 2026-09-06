import { toMarkdownBytes, formatFromExtension } from '@firecrawl/anydoc';
process.once('message', async ({ data, format }) => {
  try {
    const text = await toMarkdownBytes(Buffer.from(data, 'base64'), formatFromExtension(format));
    process.send({ ok: true, text }, () => process.exit(0));
  } catch (error) {
    process.send({ ok: false, code: error.code || 'malformed' }, () => process.exit(0));
  }
});
