#!/usr/bin/env node
/*
 * export_script.js : sort de deck.json le script à dire, la liste des apparitions et les sous-titres.
 *
 *   node export_script.js deck.json --md script.md      script complet, slide par slide (Markdown)
 *   node export_script.js deck.json --etapes            apparitions, avec l'effet conseillé par logiciel
 *   node export_script.js deck.json --srt sortie.srt    sous-titres, une phrase (ou un morceau) par sous-titre
 *
 * Les options se combinent. Aucune dépendance (Node 18 et plus).
 * Champs lus : notes (texte, liste de textes ou de segments), build, duree_s, intervenant, bloc,
 * textes, image, source, a_fournir. Format complet : ../references/deck-spec.md.
 *
 * Marqueurs : [À VALIDER : texte] (proposition de la skill) reste signalé dans le script et
 * listé à la fin ; [À FOURNIR : libellé] (ce qui manque) aussi, avec la consigne de a_fournir.
 * Aucun terme de travail (numéro d'étape, archétype) n'apparaît dans les sorties.
 */
'use strict';

const fs = require('fs');
const path = require('path');

const FOURNIR_RE = /\[(?:À|A) FOURNIR[^\]]*\]|\[TO PROVIDE[^\]]*\]/gi;
const VALIDER_RE = /\[(?:(?:À|A) VALIDER|TO VALIDATE|TO CONFIRM)[^\]]*\]/gi;
const FOURNIR_ONE = /\[(?:À|A) FOURNIR[^\]]*\]|\[TO PROVIDE[^\]]*\]/i;

// Effet conseillé par type de slide (Keynote / PowerPoint, voir references/formats/keynote.md §4).
// Le .pptx porte déjà un fondu au clic ; ces effets servent à retoucher, ou pour Keynote.
const EFFETS = {
  A1: ['Dissolution', 'Fondu'], A4: ['Dissolution, par paragraphe', 'Fondu, par paragraphe'],
  A5: ['Dissolution, puis Apparition', 'Fondu, puis Apparaître'],
  A6: ['Dissolution, élément par élément', 'Fondu, élément par élément'],
  A7: ['Dissolution, par paragraphe', 'Fondu, par paragraphe'],
  A8: ['Apparition, par paragraphe', 'Apparaître, par paragraphe'],
  A9: ['Dissolution', 'Fondu'], A11: ['Apparition', 'Apparaître'], A14: ['Dissolution', 'Fondu'],
  A19: ['Apparition', 'Apparaître'],
  A20: ['deux slides + Déplacement magique', 'deux slides + Morphose']
};

function arr(v) {
  if (v === undefined || v === null) return [];
  return Array.isArray(v) ? v : [v];
}
function str(v) {
  return v === undefined || v === null ? '' : String(v);
}
function uniq(list) {
  const out = [];
  list.forEach((x) => { if (x && !out.includes(x)) out.push(x); });
  return out;
}
function mmss(sec) {
  sec = Math.max(0, Math.round(sec));
  return Math.floor(sec / 60) + ':' + String(sec % 60).padStart(2, '0');
}
function srtTime(sec) {
  const ms = Math.max(0, Math.round(sec * 1000));
  const h = Math.floor(ms / 3600000);
  const m = Math.floor(ms / 60000) % 60;
  const s = Math.floor(ms / 1000) % 60;
  const r = ms % 1000;
  return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0') + ':' + String(s).padStart(2, '0') + ',' + String(r).padStart(3, '0');
}
function inner(m) {
  const s = String(m).slice(1, -1);
  const i = s.indexOf(':');
  return i < 0 ? '' : s.slice(i + 1).trim();
}
function propsOf(s) {
  return (str(s).match(VALIDER_RE) || []).map(inner).filter(Boolean);
}
function provideOf(s) {
  return (str(s).match(FOURNIR_RE) || []).map(inner).filter(Boolean);
}
function unmark(s) {
  return str(s).replace(VALIDER_RE, (m) => inner(m));
}

// Segments de parole d'une slide : [{ intervenant, texte, type }]
function segments(slide, def) {
  const n = slide.notes;
  if (n === undefined || n === null || n === '') return [];
  const who = slide.intervenant || def;
  return arr(n)
    .map((s) => (typeof s === 'string' ? { intervenant: who, texte: s } : { intervenant: s.intervenant || who, texte: String(s.texte || s.text || ''), type: s.type }))
    .filter((s) => s.texte.trim() !== '');
}
function buildList(slide) {
  return arr(slide.build).map((b) => (typeof b === 'string' ? b : str(b && (b.texte || b.text)))).filter((b) => b.trim());
}

// Textes affichés, dans l'ordre de lecture, en une ligne.
function screenText(slide) {
  const t = slide.textes || {};
  const out = [];
  const add = (v) => { if (str(v).trim()) out.push(str(v).replace(/\n/g, ' / ').replace(/\*([^*]+)\*/g, '$1').trim()); };
  add(t.surtitre);
  add(t.libelle && !t.chiffre ? t.libelle : '');
  add(t.titre);
  add(t.chiffre);
  if (t.chiffre) add(t.libelle);
  arr(t.lignes).forEach(add);
  add(t.citation);
  add(t.auteur);
  arr(t.offres).forEach((o) => add([o.nom, o.mention, o.prix, o.note].filter(Boolean).join(' ')));
  arr(t.tuiles).forEach((x) => add(typeof x === 'string' ? x : [x.chiffre, x.titre, x.texte].filter(Boolean).join(' ')));
  arr(t.barres).forEach((b) => add([b.libelle, b.etiquette].filter(Boolean).join(' ')));
  arr(t.logos).forEach((l) => add(typeof l === 'string' ? l : l.nom));
  arr(t.elements).forEach((e) => add(typeof e === 'string' ? e : e.titre));
  add(t.nom);
  add(t.fonction);
  add(t.note);
  add(t.axe);
  return out.join(' · ');
}
function screenStrings(slide) {
  const all = [];
  const walk = (v) => {
    if (typeof v === 'string') all.push(v);
    else if (Array.isArray(v)) v.forEach(walk);
    else if (v && typeof v === 'object') Object.keys(v).forEach((k) => { if (k !== 'image') walk(v[k]); });
  };
  walk(slide.textes || {});
  if (slide.source) all.push(String(slide.source));
  return all;
}
// Images qui manquent encore (chemin introuvable), relatives à deck.json.
function missingImages(slide, deckDir) {
  const out = [];
  const check = (spec) => {
    if (!spec) return;
    const s = typeof spec === 'string' ? { chemin: spec } : spec;
    const abs = s.chemin ? path.resolve(deckDir, s.chemin) : '';
    if (!abs || !fs.existsSync(abs)) out.push(s);
  };
  check(slide.image);
  check(slide.image_fond);
  const t = slide.textes || {};
  arr(t.tuiles).forEach((x) => { if (x && typeof x === 'object' && x.image) check(x.image); });
  arr(t.elements).forEach((x) => { if (x && typeof x === 'object' && x.image) check(x.image); });
  arr(t.logos).forEach((x) => { if (x && typeof x === 'object' && x.image) check({ chemin: x.image, attendu: x.nom }); });
  return out;
}

// Coupe un texte en sous-titres de 2 lignes de 42 signes au plus.
function subtitleChunks(text) {
  const sentences = text.replace(/\s+/g, ' ').trim().match(/[^.!?…]+[.!?…]*["»”)]?\s*/g) || [];
  const chunks = [];
  const wrap = (sent, width) => {
    const lines = [];
    let cur = '';
    sent.split(' ').forEach((w) => {
      if (cur && (cur + ' ' + w).length > width) {
        lines.push(cur);
        cur = w;
      } else {
        cur = cur ? cur + ' ' + w : w;
      }
    });
    if (cur) lines.push(cur);
    return lines;
  };
  sentences.forEach((sent) => {
    sent = sent.trim();
    // Lignes de 42 signes au plus, deux lignes par sous-titre ; si la dernière resterait seule,
    // on répartit la phrase sur un nombre pair de lignes plus courtes.
    let lines = wrap(sent, 42);
    if (lines.length > 1 && lines.length % 2 === 1) {
      const even = wrap(sent, Math.max(20, Math.ceil(sent.length / (lines.length + 1)) + 6));
      if (even.length % 2 === 0 && even.every((l) => l.length <= 42)) lines = even;
    }
    for (let i = 0; i < lines.length; i += 2) chunks.push(lines.slice(i, i + 2).join('\n'));
  });
  return chunks;
}

function main() {
  const args = process.argv.slice(2);
  if (!args.length || args.includes('-h') || args.includes('--help')) {
    console.log('Usage : node export_script.js deck.json [--md script.md] [--etapes] [--srt sortie.srt]');
    process.exit(args.length ? 0 : 2);
  }
  let deckPath = null;
  let md = null;
  let srt = null;
  let etapes = false;
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--md') md = args[++i];
    else if (args[i] === '--srt') srt = args[++i];
    else if (args[i] === '--etapes') etapes = true;
    else deckPath = args[i];
  }
  if (!deckPath || (!md && !srt && !etapes)) {
    console.error('Indiquez deck.json et au moins une sortie : --md, --etapes ou --srt.');
    process.exit(2);
  }
  let deck;
  try {
    deck = JSON.parse(fs.readFileSync(deckPath, 'utf8'));
  } catch (e) {
    console.error('Impossible de lire ' + deckPath + ' : ' + e.message);
    process.exit(2);
  }
  const deckDir = path.dirname(path.resolve(deckPath));
  const meta = deck.meta || {};
  const en = String(meta.langue || 'fr').toLowerCase().startsWith('en');
  const slides = arr(deck.slides);
  const people = {};
  arr(deck.intervenants).forEach((p) => { if (p && p.id) people[p.id] = p; });
  const def = (arr(deck.intervenants)[0] || {}).id || '';
  const T = en
    ? {
      sep: ': ', relay: 'handover', build: 'Build', source: 'Source', screen: 'On screen', validate: 'To validate', provide: 'To provide',
      total: 'Total duration', speakers: 'Speakers', none: 'none', roleMissing: 'role to provide', register: 'register',
      withAudience: ' with the audience, ', between: ' between speakers', forTarget: ' for ', onScreen: 'on screen', inScript: 'in the script',
      image: 'screenshot or image', q: (s) => '“' + s + '”', markersInScript: (n) => n + ' marker(s) in the script',
      etapesHead: 'The .pptx already fades each slide in, and fades these builds in on click. This list is for retouching, Keynote, Canva, Figma Slides or video.',
      noBuild: 'No build steps in deck.json.'
    }
    : {
      sep: ' : ', relay: 'relais', build: 'Apparitions', source: 'Source', screen: 'À l\'écran', validate: 'À valider', provide: 'Reste à fournir',
      total: 'Durée totale', speakers: 'Intervenants', none: 'aucun', roleMissing: 'fonction à fournir', register: 'registre',
      withAudience: ' au public, ', between: ' entre intervenants', forTarget: ' pour ', onScreen: 'à l\'écran', inScript: 'dans le script',
      image: 'capture ou image', q: (s) => '« ' + s + ' »', markersInScript: (n) => n + ' marqueur(s) dans le script',
      etapesHead: 'Le .pptx porte déjà un fondu entre les slides et ces apparitions en fondu, au clic. Cette liste sert à les retoucher, ou pour Keynote, Canva, Figma Slides et la vidéo.',
      noBuild: 'Aucune apparition prévue dans deck.json.'
    };
  const roleOf = (p) => {
    const r = str(p.role).trim();
    return r && !FOURNIR_ONE.test(r) ? unmark(r) : '';
  };

  // Script complet (Markdown)
  if (md) {
    let t = 0;
    const toValidate = [];
    const toProvide = [];
    let o = '# ' + unmark(meta.titre || 'Script') + '\n\n';
    o += T.speakers + T.sep + arr(deck.intervenants).map((p) => {
      const r = roleOf(p);
      return unmark(p.nom || p.id) + ' (' + (r || T.roleMissing) + ')';
    }).join(', ');
    if (meta.registre) o += ' · ' + T.register + T.sep + meta.registre + (meta.registre_intervenants ? T.withAudience + meta.registre_intervenants + T.between : '');
    o += '\n';
    arr(deck.intervenants).forEach((p) => {
      if (!roleOf(p)) toProvide.push('- ' + T.speakers + T.sep + unmark(p.nom || p.id) + ', ' + T.roleMissing + '.');
    });
    slides.forEach((s, i) => {
      const d = Number(s.duree_s) || 0;
      o += '\n## ' + (i + 1) + '. ' + (s.bloc ? s.bloc + ' · ' : '') + mmss(t) + ' → ' + mmss(t + d) + ' (' + d + ' s)\n';
      const scr = screenText(s);
      if (scr) o += '\n_' + T.screen + '_' + T.sep + scr + '\n';
      let last = null;
      const segs = segments(s, def);
      segs.forEach((g) => {
        const p = people[g.intervenant] || {};
        const name = unmark(p.nom || g.intervenant || '?');
        o += '\n' + (last && last !== g.intervenant ? '→ ' : '') + '**' + name + '**' + (g.type === 'relais' ? ' (' + T.relay + ')' : '') + T.sep + g.texte.trim() + '\n';
        last = g.intervenant;
      });
      const b = buildList(s);
      if (b.length) o += '\n_' + T.build + '_' + T.sep + b.join(' → ') + '\n';
      if (s.source) o += '\n_' + T.source + '_' + T.sep + s.source + '\n';
      // À valider : propositions affichées et dites.
      const onScreen = uniq([].concat.apply([], screenStrings(s).map(propsOf)));
      const said = uniq([].concat.apply([], segs.map((g) => propsOf(g.texte)))).filter((x) => !onScreen.includes(x));
      if (onScreen.length || said.length) {
        const parts = [];
        const semi = en ? '; ' : ' ; ';
        if (onScreen.length) parts.push(T.onScreen + ' ' + onScreen.map(T.q).join(semi));
        if (said.length) parts.push(T.inScript + ' ' + said.map(T.q).join(semi));
        toValidate.push('- ' + (i + 1) + '. ' + (s.id || '') + T.sep + parts.join(' · '));
      }
      // Reste à fournir : libellés à l'écran, consignes, images, slide AF, marqueurs du script.
      const prov = [];
      const labels = uniq([].concat.apply([], screenStrings(s).map(provideOf)));
      if (labels.length) prov.push(labels.join(', ') + '.');
      arr(s.a_fournir).forEach((x) => { if (str(x).trim()) prov.push(str(x).trim()); });
      if (String(s.archetype || '').toUpperCase() === 'AF' && (s.textes || {}).attendu) prov.push(str(s.textes.attendu).trim());
      missingImages(s, deckDir).forEach((im) => prov.push(T.image + (im.attendu ? ' ' + T.q(im.attendu) : '') + (im.chemin ? ' (' + im.chemin + ')' : '') + '.'));
      const inScript = segs.reduce((n, g) => n + provideOf(g.texte).length, 0);
      if (inScript) prov.push(T.markersInScript(inScript) + '.');
      if (prov.length) {
        o += '\n_' + T.provide + '_' + T.sep + uniq(prov).join(' ') + '\n';
        toProvide.push('- ' + (i + 1) + '. ' + (s.id || '') + T.sep + uniq(prov).join(' '));
      }
      t += d;
    });
    o += '\n---\n\n' + T.total + T.sep + mmss(t) + (meta.duree_min ? T.forTarget + meta.duree_min + ' min' : '') + '\n';
    o += '\n' + T.validate + T.sep + (toValidate.length ? '\n\n' + toValidate.join('\n') : T.none) + '\n';
    o += '\n' + T.provide + T.sep + (toProvide.length ? '\n\n' + toProvide.join('\n') : T.none) + '\n';
    fs.writeFileSync(md, o);
    console.log('Script → ' + path.resolve(md) + ' (' + slides.length + ' slides, ' + mmss(t) + ' ; ' + toValidate.length + ' slide(s) à valider, ' + toProvide.length + ' ligne(s) à fournir)');
  }

  // Apparitions : la liste de build, avec l'effet conseillé par logiciel.
  if (etapes) {
    let n = 0;
    console.log(T.etapesHead);
    slides.forEach((s, i) => {
      const b = buildList(s);
      if (!b.length) return;
      n += 1;
      const arch = String(s.archetype || '').toUpperCase();
      const e = EFFETS[arch];
      const effet = e ? ' [Keynote : ' + e[0] + ' · PowerPoint : ' + e[1] + ']' : '';
      console.log((i + 1) + '. ' + (s.id || '') + effet + '\n   ' + b.map((x, k) => (k + 1) + ') ' + x).join('\n   '));
    });
    if (!n) console.log(T.noBuild);
  }

  // Sous-titres : le texte dit, propositions comprises (sans leur marqueur).
  if (srt) {
    let t = 0;
    let k = 0;
    let o = '';
    let markers = 0;
    slides.forEach((s) => {
      const d = Number(s.duree_s) || 0;
      const text = unmark(segments(s, def).map((g) => g.texte.trim()).join(' '));
      markers += (text.match(FOURNIR_RE) || []).length;
      const chunks = subtitleChunks(text);
      const total = chunks.reduce((a, c) => a + c.length, 0) || 1;
      let u = t;
      chunks.forEach((c) => {
        const dt = d * c.length / total;
        k += 1;
        o += k + '\n' + srtTime(u) + ' --> ' + srtTime(u + dt) + '\n' + c + '\n\n';
        u += dt;
      });
      t += d;
    });
    fs.writeFileSync(srt, o);
    console.log('Sous-titres → ' + path.resolve(srt) + ' (' + k + ' sous-titres, ' + mmss(t) + ')');
    if (markers) console.log('Attention : ' + markers + ' marqueur(s) [À FOURNIR] dans le texte à dire.');
  }
}

main();
