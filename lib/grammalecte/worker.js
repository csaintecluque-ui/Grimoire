// Grimoire — correcteur Grammalecte dans un Web Worker (l'écran ne fige pas pendant l'analyse).
// Messages reçus : { id, kind: 'check', paras: [texte…], opts? } → { id, out: [{ g: [fautes], s: [mots inconnus] }…] }
//                  { id, kind: 'suggest', word }                 → { id, out: [suggestions…] }
'use strict';
self.window = self;
self.GL_BASE = new URL('./', self.location.href).href;
importScripts(
  'graphspell/helpers.js', 'graphspell/lexgraph_fr.js', 'graphspell/char_player.js',
  'fr/conj.js', 'fr/mfsp.js', 'fr/phonet.js', 'fr/gc_options.js', 'fr/gc_engine.js', 'fr/gc_functions.js',
  'fr/gc_rules.js', 'fr/gc_rules_graph.js', 'fr/cregex.js', 'text.js',
  'graphspell/spellchecker.js', 'graphspell/tokenizer.js', 'graphspell/str_transform.js', 'graphspell/ibdawg.js'
);
let gce = null, sp = null;
function init() {
  gce = self.grammalecte.gc_engine;
  gce.load('JavaScript', 'aRGB');
  sp = gce.getSpellChecker();
  if (!sp) throw new Error('dictionnaire introuvable');
}
self.onmessage = e => {
  const d = e.data || {};
  try {
    if (!gce) init();
    if (d.kind === 'check') {
      if (d.opts) gce.setOptions(new Map(Object.entries(d.opts)));
      const out = (d.paras || []).map(t => {
        if (!t || !/\S/.test(t)) return { g: [], s: [] };
        let g = [];
        try { g = gce.parse(t, 'FR', false, null, false) || []; } catch (err) { g = []; }
        let s = [];
        try { s = sp.parseParagraph(t) || []; } catch (err) { s = []; }
        return {
          g: g.map(x => ({ a: x.nStart, b: x.nEnd, msg: x.sMessage, sug: (x.aSuggestions || []).slice(0, 8), type: x.sType, rule: x.sRuleId, url: x.URL || '' })),
          s: s.map(x => ({ a: x.nStart, b: x.nEnd, w: x.sValue }))
        };
      });
      self.postMessage({ id: d.id, out });
    } else if (d.kind === 'suggest') {
      const list = [];
      for (const l of sp.suggest(String(d.word || ''), 8)) for (const w of l) if (!list.includes(w)) list.push(w);
      self.postMessage({ id: d.id, out: list.slice(0, 8) });
    } else self.postMessage({ id: d.id, out: { version: gce.version } });
  } catch (err) {
    self.postMessage({ id: d.id, err: String(err && err.message || err) });
  }
};
