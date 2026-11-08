// Grimoire — qpdf (Apache 2.0, compilé en WebAssembly) dans un Web Worker :
// chiffrement et déchiffrement des PDF protégés par mot de passe, sans figer l'écran.
// Message reçu : { id, args: [...], files: { '/in.pdf': Uint8Array, ... }, out: '/out.pdf' }
// Réponse     : { id, code, out: Uint8Array|null, stdout, stderr }
'use strict';
importScripts('qpdf.js');
const factory = self.Module;
let compiled = null;
function wasmModule() {
  if (!compiled) compiled = WebAssembly.compileStreaming ? WebAssembly.compileStreaming(fetch('qpdf.wasm')) : fetch('qpdf.wasm').then(r => r.arrayBuffer()).then(b => WebAssembly.compile(b));
  return compiled;
}
async function run(d) {
  const stdout = [], stderr = [];
  // cette version de qpdf écrit dans console.log / console.error, liés à la création de l'instance :
  // on les remplace le temps de la créer, pour récupérer ce que qpdf affiche
  const log = console.log, err = console.error;
  try {
    const mod = await wasmModule();
    console.log = (...a) => stdout.push(a.join(' ')); console.error = (...a) => stderr.push(a.join(' '));
    let q;
    try {
      // une instance neuve par opération : qpdf « quitte » à la fin de chaque commande
      q = await factory({ noInitialRun: true, instantiateWasm: (imports, ok) => { WebAssembly.instantiate(mod, imports).then(inst => ok(inst, mod)); return {}; } });
    } finally { console.log = log; console.error = err; }
    for (const [p, b] of Object.entries(d.files || {})) q.FS.writeFile(p, b);
    let code;
    try { code = q.callMain(d.args || []); } catch (e) { code = e && typeof e.status === 'number' ? e.status : -1; if (code === -1) stderr.push(String(e && e.message || e)); }
    let out = null;
    if (d.out) { try { out = q.FS.readFile(d.out); } catch {} }
    return { id: d.id, code, out, stdout: stdout.join('\n'), stderr: stderr.join('\n') };
  } catch (e) {
    console.log = log; console.error = err;
    return { id: d.id, code: -1, out: null, stdout: stdout.join('\n'), stderr: String(e && e.message || e) };
  }
}
// une opération à la fois
let queue = Promise.resolve();
self.onmessage = e => {
  const d = e.data || {};
  queue = queue.then(() => run(d)).then(r => self.postMessage(r, r.out ? [r.out.buffer] : []));
};
