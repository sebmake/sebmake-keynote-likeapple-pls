#!/usr/bin/env node
/*
 * build_deck.js : rend un deck.json en présentation 16:9 (.pptx) et, en option, en aperçu HTML.
 *
 *   node build_deck.js deck.json sortie.pptx [--preview sortie.html] [--polices auto|repli]
 *
 * Polices : le .pptx écrit toujours celles du thème (nom exact), même absentes de cette machine ;
 * l'aperçu prend alors la police de repli, et la console le signale. --polices repli écrit le repli.
 * Le .pptx s'ouvre dans PowerPoint et dans Keynote, et s'importe dans Canva. Chaque slide entre
 * par un fondu, et les éléments décrits par slide.build apparaissent en fondu, au clic.
 * L'aperçu HTML reprend la même mise en page (mêmes coordonnées, mêmes tailles), sans les
 * animations. Ajoutez #qa à son adresse pour entourer en rouge les textes qui débordent.
 *
 * Tous les réglages visuels (corps, graisses, suivis, interlignes, positions, couleurs, durées) sont
 * dans STYLE, plus bas : ce sont les mesures de ../references/style-visuel.md.
 * Format de deck.json : voir ../references/deck-spec.md.
 */
'use strict';

const fs = require('fs');
const path = require('path');

let PptxGenJS;
try {
  PptxGenJS = require('pptxgenjs');
} catch (e) {
  console.error('pptxgenjs est introuvable. Lancez « npm install » dans ' + __dirname + ' puis relancez.');
  process.exit(2);
}
function loadJSZip() {
  try {
    return require(require.resolve('jszip', { paths: [path.dirname(require.resolve('pptxgenjs')), __dirname] }));
  } catch (e) {
    return null;
  }
}

// ===========================================================================
// STYLE : tous les réglages visuels du moteur, tirés de ../references/style-visuel.md (mesures
// relevées sur les keynotes du corpus, converties pour le canevas de 10 × 5,625 in).
// - Positions en pouces. « cy » est le centre visé de la zone des capitales d'un bloc : la boîte
//   est posée par son centre, en alignement vertical « milieu » (style-visuel §0).
// - Corps en points, donnés pour une police dont la capitale fait 0,728 em (Inter) : le moteur les
//   ajuste à la capitale de la police écrite, pour garder la même hauteur de capitale.
// - [corps maximal, plancher] : le moteur réduit jusqu'à ce que le texte tienne, sans descendre
//   sous le plancher ; au-delà, il le signale : il faut couper des mots, pas réduire encore.
// ===========================================================================
const STYLE = {
  // Par police : capitale par em (§0), suivi en em pour [≥ 50 pt, 28 à 50 pt, < 28 pt] (§2.2), et
  // facteur de l'estimation de largeur (étalonnée sur Arial, prudente pour les autres polices).
  polices: {
    'inter tight': { cap: 0.728, suivi: [-0.02, -0.02, 0], largeur: 0.96 },
    inter: { cap: 0.728, suivi: [-0.035, -0.035, -0.02], largeur: 1 },
    'helvetica neue': { cap: 0.714, suivi: [-0.045, -0.03, -0.01], largeur: 0.97 },
    helvetica: { cap: 0.717, suivi: [-0.045, -0.03, -0.01], largeur: 0.97 },
    arial: { cap: 0.716, suivi: [-0.06, -0.045, -0.015], largeur: 0.93 },
    defaut: { cap: 0.72, suivi: [-0.03, -0.025, -0.01], largeur: 1 }
  },
  capReference: 0.728,
  suivi: { max: -0.07, sansSerrageSous: 14 },       // jamais au-delà de −0,07 em, aucun serrage sous 14 pt
  // Interligne en multiple de la cap (§2.3) : interligne (pt) = k × cap par em × corps. Les listes
  // gardent l'interligne des titres et posent le reste du pas en espace après le paragraphe.
  interligne: { titre: 1.6, citation: 1.5, sommaire: 1.4, objet: 1.5, liste: 1.95, liste2: 3.1, courte: 1.8, texte: 1.6 },
  // Graisse (§2.1) : demi-gras (face « SemiBold » installée, sinon le gras de la famille) pour les
  // titres, noms, chiffres, listes, citations et prix ; regular pour le gris et les mentions.
  graisse: { texte: 'regular' },
  // Mélanges de couleur (§2.4) : texte secondaire à 53 % du texte vers le fond ; l'autre barre d'un
  // graphique, gris à 45 % ; tuiles de bento sur fond sombre, texte mélangé à 12 %.
  couleur: { secondaire: 0.53, autre: 0.55, tuileSombre: 0.12 },
  // Surtitre : corps = 0,42 × le titre, entre 16 et 28 pt ; ligne de base à 0,45 cap au-dessus du
  // haut des capitales du titre. En accent et demi-gras si l'accent de la slide est libre, sinon gris.
  surtitre: { rapport: 0.42, min: 16, max: 28, ecart: 0.45 },
  A1: { corps: [100, 64], uneLigne: 6.4, large: [64, 56], deuxLignes: [72, 60], x: 0.6, w: 8.8, cy: 2.81 },
  // Slide finale (un mot) : plus gros, centré à 47,5 % H ; en pleine page d'accent, la seule du deck,
  // si son texte (fond_sombre ou blanc) s'y lit à 4,5:1.
  finale: { corps: [120, 64], cy: 2.67, pleinePageAccent: true, contraste: 4.5 },
  A2: { icone: 1.9, nom: [56, 32], ecart: 0.42, cy: 2.81 },
  A3: { corps: { 1: 80, 2: 80, 3: 80, 4: 68, 5: 56, 6: 44, 7: 40, 8: 36 }, plancher: 28, x: 1.0, w: 8.0, cy: 2.75 },
  A4: { corps: [50, 44], x: 1.2, w: 7.6, cy: 2.73 },
  A5: { chiffre: [144, 110], largeur: 8.0, cy: 2.81, cyDeuxLignes: 2.56, libelle: [24, 18], libelleCy: 4.5, libelleHautDeuxLignes: 3.54, libelleX: 0.8, libelleW: 8.4, surtitre: 24, surtitreEcart: 0.25 },
  A6: { titre: [28, 22], titreHaut: 0.5, titreHautSurtitre: 0.8, zone: { x: 1.5, y: 1.9, w: 7.0, h: 2.4 }, barre: 0.06, etiquette: [23, 14], axe: 16 },
  A7: { lignes: [36, 28], x: 1.0, w: 8.0, cy: 2.64, titre: 24 },
  A8: { lignes: [34, 24], picto: 1.12, ecart: 0.75, largeurMax: 7.4, cy: 2.64, titre: 24 },
  A9: { objet: 5.6, colonne: { x: 6.3, w: 3.4 }, titre: [34, 26], lignes: 24, ecartLignes: 0.28, cy: 2.81 },
  A10: { capture: 0.85, appareil: 5.2, rayon: 0.12 },
  A13: { h: 4.6, w: 8.0 },
  A14: { marge: 0.09, gouttiere: 0.085, rayon: 0.2, centre: { w: 3.9, h: 2.25 }, nom: [60, 36], titre: [24, 15], libelle: 12, chiffre: [48, 28], pad: 0.17 },
  A15: { libelle: [11, 8] },
  A16: { titre: 24, logoH: 0.56, rangee: 1.1, largeur: 8.0, nom: [20, 14] },
  A17: { citation: [36, 28], x: 0.92, w: 7.6, auteur: [26, 18], ecart: 0.53, cy: 2.81 },
  A18: { titre: [42, 28], note: [11, 9] },
  A19: { nom: [40, 28], petit: 16, prix: [36, 28], ecartNomHaut: 0.48, ecartNomColonne: 0.25, ecartPetit: 0.16, ecartOffres: 0.79, nomHaut: 0.72, objetHaut: { y: 2.7, bas: 4.9, w: 6.0 }, colonne: { cx: 7.0, w: 4.2 }, objetColonne: { cx: 2.5, w: 2.8, y: 0.26, h: 5.1 }, deuxColonnes: [2.75, 7.25], heros: [144, 110], periode: 18, cy: 2.81,
    // Deux offres sans visuel (convention) : la pile garde ses rapports mais grandit pour tenir la place.
    sansVisuel: { nom: [56, 36], petit: 20, prix: [52, 32], ecartPetit: 0.24 } },
  A20: { libelle: 16, libelleHaut: 0.45, mot: [40, 30], ecart: 0.19, visuel: { y: 1.6, bas: 5.0, w: 9.2 }, seul: { libelle: 24, ecart: 0.47 } },
  A21: { corps: 16, x: 0.54, bases: [4.75, 5.08] },
  AF: { titre: [50, 36] },
  SOURCES: { titre: [20, 14], liste: [13, 8] },
  source: [11, 9],
  // [À FOURNIR] : un emplacement, pas du contenu. Son libellé est estompé à toutes les tailles :
  // couleur du texte qu'il remplace, à « texte » d'opacité (alpha du run dans le .pptx), jamais en
  // dégradé ; souligné en pointillé d'accent, discret (« opacite »). Le soulignement d'une police
  // grossit avec le corps (Arial : 0,07 em, soit 10 pt de points sous un chiffre de 144 pt) : au-dessus
  // de « seuil » pt, le moteur le remplace par un filet pointillé d'épaisseur fixe, posé sous la ligne
  // de base (« ecart » em, borné en points), le même dans le .pptx et l'aperçu, chacun placé pour sa police.
  // « chasse » : [demi-gras ou gras, regular], rapport réel / estimé de la largeur du marqueur, mesuré dans
  // l'aperçu avec le facteur de la police affichée.
  marqueur: { texte: 0.4, trait: 'dotted', opacite: 0.7, seuil: 26, epaisseur: 1.5, ecart: 0.085, ecartPt: [3, 12], chasse: [0.99, 0.99] },
  // Interface schématique d'une image qui manque (§5.2) : carte sans trait, barres grises à bouts
  // ronds, longueurs 100 / 72 / 48 %, une tous les 0,16 in, une seule ligne en accent.
  // Carte claire sur fond sombre, grise à 5 % sur fond blanc ; accent en pastille sous 1,8:1.
  schema: { rayon: 0.12, barre: 0.08, pas: 0.16, longueurs: [1, 0.72, 0.48], carteGrise: 0.05, barres: [0.22, 0.38], contrasteBarre: 1.8, ombre: { flou: 0.3, decalage: 0.08, opacite: 0.12 }, appareil: { coin: 0.13, bord: 0.03 } },
  // Équilibrage des lignes : pénalité d'une ligne qui finit sur un mot-outil, en part du carré de
  // la longueur moyenne (0,25 : une ligne une fois et demie plus longue que l'autre reste préférée).
  coupure: { motOutil: 0.25 },
  transition: { vitesse: 'fast' },                // fondu de 0,5 s entre les slides
  apparition: { duree_ms: 400, suite_ms: 600 }    // fondu d'entrée ; délai d'une apparition « qui suit »
};

// ---------------------------------------------------------------------------
// Constantes de mise en page (pouces). LAYOUT_16x9 = 10 x 5,625.
// ---------------------------------------------------------------------------
const W = 10;
const H = 5.625;
const M = 0.6;
const FOURNIR_SRC = '\\[(?:À|A) FOURNIR[^\\]]*\\]|\\[TO PROVIDE[^\\]]*\\]';
const VALIDER_SRC = '\\[(?:(?:À|A) VALIDER|TO VALIDATE|TO CONFIRM)[^\\]]*\\]';
const DATA_DIR = path.join(__dirname, '..', 'references', 'data');

// Archétypes dont l'image (slide.image) est le contenu, pas le fond.
const CONTENT_IMAGE = new Set(['A2', 'A9', 'A10', 'A11', 'A13', 'A19', 'A20']);
// Archétypes plein cadre : l'image couvre toute la slide.
const FULL_BLEED = new Set(['A12', 'A18', 'A21', 'A22', 'A23', 'A24', 'A25']);

const LABELS = {
  fr: {
    build: 'APPARITIONS', timing: 'MINUTAGE', source: 'SOURCE', relay: 'relais', sep: ' : ',
    toValidate: 'À VALIDER', toProvide: 'À FOURNIR', role: 'fonction',
    sourceShown: 'affichée sur la slide', sourceFinal: 'reportée au carton final', sourceHidden: 'non affichée',
    sourcesTitle: 'Sources et mentions', missingFile: 'introuvable', slideWord: 'slide',
    noNotes: '[À FOURNIR : script de cette slide]', markerDefault: 'à fournir', content: 'contenu',
    finalNotes: 'Carton final : sources et mentions. Il reste à l\'écran pendant les remerciements ; on ne le commente pas.',
    sourceRef: (n) => 'Source ' + n + ' · en fin de présentation', sourceRenvoi: 'renvoi n°',
    image: 'Image', q: (s) => '« ' + s + ' »',
    cap: { interface: 'Capture à fournir', appareil: 'Capture à fournir', objet: 'Visuel à fournir', photo: 'Photo à fournir', carte: 'Carte à fournir', icone: 'Icône à fournir', logo: 'Logo à fournir', tuile: 'Image à fournir' },
    kind: { interface: 'capture', appareil: 'capture', objet: 'visuel', photo: 'photo', carte: 'carte', icone: 'icône', logo: 'logo', tuile: 'image' },
    parts: { surtitre: 'surtitre', titre: 'titre', lignes: 'ligne', chiffre: 'chiffre', libelle: 'libellé', visuel: 'visuel', barres: 'barre', tuiles: 'tuile', centre: 'tuile centrale', prix: 'prix', nom: 'nom', date: 'date', citation: 'citation', auteur: 'auteur', logos: 'logo', note: 'note', fonction: 'fonction' }
  },
  en: {
    build: 'BUILD', timing: 'TIMING', source: 'SOURCE', relay: 'handover', sep: ': ',
    toValidate: 'TO VALIDATE', toProvide: 'TO PROVIDE', role: 'role',
    sourceShown: 'shown on the slide', sourceFinal: 'moved to the final card', sourceHidden: 'not shown',
    sourcesTitle: 'Sources and notes', missingFile: 'not found', slideWord: 'slide',
    noNotes: '[TO PROVIDE: script for this slide]', markerDefault: 'to provide', content: 'content',
    finalNotes: 'Final card: sources and legal notes. Leave it on screen during the thanks; do not comment on it.',
    sourceRef: (n) => 'Source ' + n + ' · see the final slide', sourceRenvoi: 'note no. ',
    image: 'Image', q: (s) => '“' + s + '”',
    cap: { interface: 'Screenshot to provide', appareil: 'Screenshot to provide', objet: 'Visual to provide', photo: 'Photo to provide', carte: 'Map to provide', icone: 'Icon to provide', logo: 'Logo to provide', tuile: 'Image to provide' },
    kind: { interface: 'screenshot', appareil: 'screenshot', objet: 'visual', photo: 'photo', carte: 'map', icone: 'icon', logo: 'logo', tuile: 'image' },
    parts: { surtitre: 'kicker', titre: 'title', lignes: 'line', chiffre: 'figure', libelle: 'label', visuel: 'visual', barres: 'bar', tuiles: 'tile', centre: 'centre tile', prix: 'price', nom: 'name', date: 'date', citation: 'quote', auteur: 'author', logos: 'logo', note: 'note', fonction: 'role' }
  }
};

// ---------------------------------------------------------------------------
// Petits utilitaires
// ---------------------------------------------------------------------------
function readJSON(file, fallback) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (e) {
    return fallback;
  }
}
function arr(v) {
  if (v === undefined || v === null) return [];
  return Array.isArray(v) ? v : [v];
}
function str(v) {
  return v === undefined || v === null ? '' : String(v);
}
function round2(n) {
  return Math.round(n * 1000) / 1000;
}
function clamp(v, a, b) {
  return Math.max(a, Math.min(b, v));
}
function escHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function escXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
// Nom de police écrit tel quel par pptxgenjs dans un attribut XML (typeface="…") : échappé ici, sinon
// un nom comme « Grotesk & Co » rend le .pptx illisible.
function xmlFace(s) {
  return escXml(String(s).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, ''));
}
function unescXml(s) {
  return String(s).replace(/&quot;/g, '"').replace(/&apos;/g, '\'').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}
function normKey(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
}
function fmtTime(sec) {
  sec = Math.max(0, Math.round(sec));
  const m = Math.floor(sec / 60);
  const s = sec % 60;
  return m + ':' + String(s).padStart(2, '0');
}
function uniq(list) {
  const out = [];
  list.forEach((x) => { if (x && !out.includes(x)) out.push(x); });
  return out;
}
function rng(seed) {
  let s = Math.abs(Math.floor(seed)) % 233280 || 7;
  return () => {
    s = (s * 9301 + 49297) % 233280;
    return s / 233280;
  };
}

// Couleurs ------------------------------------------------------------------
function hex(c, fallback, label, warnings) {
  if (c === undefined || c === null || c === '') return fallback;
  const s = String(c).trim().replace(/^#/, '');
  if (/^[0-9a-f]{3}$/i.test(s)) return s.split('').map((ch) => ch + ch).join('').toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(s)) return s.toUpperCase();
  if (warnings) warnings.push('Couleur « ' + label + ' » illisible (' + c + ') : valeur par défaut utilisée.');
  return fallback;
}
function rgbOf(h) {
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}
function mix(a, b, t) {
  const A = rgbOf(a);
  const B = rgbOf(b);
  return A.map((v, i) => Math.round(v + (B[i] - v) * t).toString(16).padStart(2, '0')).join('').toUpperCase();
}
function lum(h) {
  return rgbOf(h).map((v) => {
    v /= 255;
    return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
  }).reduce((acc, v, i) => acc + v * [0.2126, 0.7152, 0.0722][i], 0);
}
function contrast(a, b) {
  const la = lum(a);
  const lb = lum(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}
function readableOn(bg, a, b) {
  return contrast(bg, a) >= contrast(bg, b) ? a : b;
}
// Texte posé sur la couleur d'accent : blanc si le contraste suffit pour un grand texte gras (3:1),
// sinon le fond sombre du thème (ou un quasi-noir).
function onAccent(bg, dark) {
  if (contrast(bg, 'FFFFFF') >= 3) return 'FFFFFF';
  return dark && contrast(bg, dark) >= 3 ? dark : readableOn(bg, 'FFFFFF', '111111');
}
// Accent utilisé pour du texte : l'accent lui-même s'il se lit sur le fond, sinon l'accent
// rapproché de la couleur du texte juste assez pour se lire (3:1).
function textAccent(bg, accent, fg) {
  if (contrast(bg, accent) >= 2.6) return accent;
  for (let t = 0.1; t <= 1.0001; t += 0.1) {
    const c = mix(accent, fg, t);
    if (contrast(bg, c) >= 3) return c;
  }
  return fg;
}

// ---------------------------------------------------------------------------
// Typographie et mesure du texte (estimation prudente, commune au .pptx et à l'aperçu)
// ---------------------------------------------------------------------------
const NBSP = ' ';
function typo(s, lang) {
  s = str(s);
  // Pas de coupure entre un nombre et son unité, ni dans un nombre à milliers.
  s = s.replace(/(\d) (?=\d{3}(?!\d))/g, '$1' + NBSP);
  s = s.replace(/(\d) (?=(?:%|€|\$|£|×|x\b|°|h\b|min\b|s\b|j\b|km\b|kg\b|m\b|Go\b|To\b|GB\b|TB\b|HT\b|TTC\b|fois\b|jours?\b|mois\b|ans?\b|heures?\b|minutes?\b|secondes?\b|semaines?\b|sites?\b|clients?\b|times\b|hours?\b|days?\b|years?\b))/g, '$1' + NBSP);
  s = s.replace(/([€$£]) (?=(?:HT|TTC)\b)/g, '$1' + NBSP);
  if (lang === 'fr') {
    s = s.replace(/ ([:;!?»])/g, NBSP + '$1').replace(/« /g, '«' + NBSP);
  }
  return s;
}
function charW(ch) {
  if (ch === ' ' || ch === NBSP) return 0.28;
  if (/[0-9]/.test(ch)) return 0.58;
  if (/[.,;:'’!|ijlI1]/.test(ch)) return 0.28;
  if (/[frt]/.test(ch)) return 0.37;
  if (/[mwMW@]/.test(ch)) return 0.86;
  if (/[«»"“”()\-–—/]/.test(ch)) return 0.4;
  if (/[€$£%&#×]/.test(ch)) return 0.66;
  if (ch !== ch.toLowerCase()) return 0.7;
  return 0.55;
}
// Police de chaque rôle (titre, texte) : clé de STYLE.polices, fixée par main() d'après la police
// écrite dans le .pptx (capitale, suivi) ; « mesure » liste aussi celle de l'aperçu quand la police
// manque ici (le repli) : la largeur estimée prend la plus large des deux.
// « ici » : la clé de la police que l'aperçu affiche (le repli quand la police du thème manque).
const FONT = { titre: { cle: 'defaut', mesure: ['defaut'], ici: 'defaut' }, texte: { cle: 'defaut', mesure: ['defaut'], ici: 'defaut' } };
function pol(role) {
  const f = FONT[role] || FONT.texte;
  return STYLE.polices[f.cle] || STYLE.polices.defaut;
}
function widthFactor(role) {
  const f = FONT[role] || FONT.texte;
  return Math.max.apply(null, (f.mesure || [f.cle]).map((k) => (STYLE.polices[k] || STYLE.polices.defaut).largeur));
}
// wf : facteur de largeur imposé (celui d'une seule police) ; par défaut, le plus large des polices du rôle.
function textWidth(s, pt, heavy, track, role, wf) {
  let u = 0;
  let n = 0;
  for (const ch of s) {
    u += charW(ch);
    n += 1;
  }
  return (u * (heavy ? 1.07 : 1) * 1.04 * (wf || widthFactor(role)) + n * (track || 0)) * pt / 72;
}
function wrapInfo(s, pt, heavy, w, track, role) {
  const words = s.split(/ +/).filter((x) => x.length);
  if (!words.length) return { lines: 1, maxWord: 0, widest: 0 };
  const sp = textWidth(' ', pt, heavy, track, role);
  let lines = 1;
  let cur = 0;
  let maxWord = 0;
  let widest = 0;
  for (const wd of words) {
    const ww = textWidth(wd, pt, heavy, track, role);
    if (ww > maxWord) maxWord = ww;
    if (cur === 0) cur = ww;
    else if (cur + sp + ww <= w) cur += sp + ww;
    else {
      widest = Math.max(widest, cur);
      lines += 1;
      cur = ww;
    }
  }
  return { lines, maxWord, widest: Math.max(widest, cur) };
}
// Marqueurs ------------------------------------------------------------------
// [À VALIDER : texte] : une proposition, affichée comme du vrai contenu (marqueur retiré).
// [À FOURNIR : libellé] : ce qui manque ; à l'écran, le libellé seul, estompé, souligné en pointillé.
function markerInner(m) {
  const s = String(m).slice(1, -1);
  const i = s.indexOf(':');
  return i < 0 ? '' : s.slice(i + 1).replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, '');
}
function hasFournir(s) {
  return new RegExp(FOURNIR_SRC, 'i').test(str(s));
}
function fournirLabels(s) {
  return (str(s).match(new RegExp(FOURNIR_SRC, 'gi')) || []).map((m) => markerInner(m));
}
function stripValider(s, found) {
  return str(s).replace(new RegExp(VALIDER_SRC, 'gi'), (m) => {
    const inner = markerInner(m);
    if (found && inner && !found.includes(inner)) found.push(inner);
    return inner;
  });
}
// Texte riche : *accent* et marqueurs [À FOURNIR …]
function parseInline(text, lang) {
  const out = [];
  const re = new RegExp('(' + FOURNIR_SRC + ')|\\*([^*]+)\\*', 'gi');
  let last = 0;
  let m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push({ t: text.slice(last, m.index) });
    if (m[1]) out.push({ t: markerInner(m[1]) || (LABELS[lang] || LABELS.fr).markerDefault, marker: true });
    // Un [À FOURNIR] écrit dans un *accent* reste un marqueur (estompé, en couleur d'accent à 40 %),
    // jamais un texte en clair en accent plein.
    else if (hasFournir(m[2])) parseInline(m[2], lang).forEach((r) => { if (r.t) out.push(Object.assign(r, { accent: true })); });
    else out.push({ t: m[2], accent: true });
    last = re.lastIndex;
  }
  if (last < text.length) out.push({ t: text.slice(last) });
  if (!out.length) out.push({ t: '' });
  return out;
}
// Point final (ou ! ?) d'un titre en couleur d'accent.
function accentPoint(runs) {
  const last = runs[runs.length - 1];
  if (!last || last.marker || last.accent) return runs;
  const m = /([^.!?…])([.!?])$/.exec(last.t) || (/^[.!?]$/.test(last.t) ? [last.t, '', last.t] : null);
  if (!m || /\.\.$|…$/.test(last.t)) return runs;
  const head = last.t.slice(0, -1);
  const out = runs.slice(0, -1);
  if (head) out.push(Object.assign({}, last, { t: head }));
  out.push({ t: last.t.slice(-1), accent: true, point: true });
  return out;
}
function shortLabel(s, maxWords) {
  let t = str(s).replace(/\([^)]*\)/g, ' ').replace(/\s+/g, ' ').trim();
  t = t.split(/[;,.!?](?:\s|$)/)[0].trim();
  const words = t.split(' ').filter(Boolean);
  if (words.length > (maxWords || 7)) t = words.slice(0, maxWords || 7).join(' ') + '…';
  return t;
}

// ---------------------------------------------------------------------------
// Images : dimensions lues dans l'en-tête (PNG, JPEG, GIF), sans dépendance
// ---------------------------------------------------------------------------
function imageInfo(file) {
  let buf;
  try {
    buf = fs.readFileSync(file);
  } catch (e) {
    return null;
  }
  if (buf.length > 24 && buf.readUInt32BE(0) === 0x89504e47) {
    return { type: 'png', mime: 'image/png', w: buf.readUInt32BE(16), h: buf.readUInt32BE(20), buf };
  }
  if (buf.length > 4 && buf[0] === 0xff && buf[1] === 0xd8) {
    let o = 2;
    while (o + 9 < buf.length) {
      if (buf[o] !== 0xff) {
        o += 1;
        continue;
      }
      const mk = buf[o + 1];
      if (mk === 0xff) {
        o += 1;
        continue;
      }
      if (mk === 0xd8 || mk === 0x01 || (mk >= 0xd0 && mk <= 0xd7)) {
        o += 2;
        continue;
      }
      if (mk >= 0xc0 && mk <= 0xcf && mk !== 0xc4 && mk !== 0xc8 && mk !== 0xcc) {
        return { type: 'jpeg', mime: 'image/jpeg', h: buf.readUInt16BE(o + 5), w: buf.readUInt16BE(o + 7), buf };
      }
      o += 2 + buf.readUInt16BE(o + 2);
    }
    return null;
  }
  if (buf.length > 10 && buf.slice(0, 3).toString('ascii') === 'GIF') {
    return { type: 'gif', mime: 'image/gif', w: buf.readUInt16LE(6), h: buf.readUInt16LE(8), buf };
  }
  return null;
}

// ---------------------------------------------------------------------------
// Polices : celles du thème dans le .pptx ; installées ici ou non, pour l'aperçu et la mesure
// ---------------------------------------------------------------------------
let fontFileIndex = null;
function installedFontNames() {
  if (fontFileIndex) return fontFileIndex;
  const home = process.env.HOME || process.env.USERPROFILE || '';
  const dirs = [
    path.join(home, 'Library', 'Fonts'), '/Library/Fonts', '/System/Library/Fonts', '/System/Library/Fonts/Supplemental',
    'C:\\Windows\\Fonts', path.join(process.env.LOCALAPPDATA || '', 'Microsoft', 'Windows', 'Fonts'),
    '/usr/share/fonts', '/usr/local/share/fonts', path.join(home, '.fonts'), path.join(home, '.local', 'share', 'fonts')
  ];
  const names = [];
  const walk = (d, depth) => {
    let list;
    try {
      list = fs.readdirSync(d, { withFileTypes: true });
    } catch (e) {
      return;
    }
    for (const ent of list) {
      if (ent.isDirectory()) {
        if (depth < 3) walk(path.join(d, ent.name), depth + 1);
      } else if (/\.(ttf|otf|ttc|dfont|woff2?)$/i.test(ent.name)) {
        names.push(ent.name.toLowerCase().replace(/[^a-z0-9]/g, ''));
      }
    }
  };
  dirs.forEach((d) => walk(d, 0));
  fontFileIndex = names;
  return names;
}
function fontInstalled(name) {
  const key = String(name).toLowerCase().replace(/[^a-z0-9]/g, '');
  if (!key) return false;
  return installedFontNames().some((f) => f.startsWith(key));
}
function fontSpec(v, fallbackName) {
  if (!v) return { nom: fallbackName, repli: fallbackName };
  if (typeof v === 'string') return { nom: v, repli: fallbackName };
  return { nom: v.nom || v.name || fallbackName, repli: v.repli || v.fallback || fallbackName };
}

// Clé de mesure d'une police (STYLE.polices) : capitale, suivi et largeur.
function fontKey(name) {
  const k = String(name || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  if (k.startsWith('inter tight')) return 'inter tight';
  if (k.startsWith('inter')) return 'inter';
  if (k.startsWith('helvetica neue')) return 'helvetica neue';
  if (k.startsWith('helvetica')) return 'helvetica';
  if (k.startsWith('arial')) return 'arial';
  return 'defaut';
}
// Face demi-grasse installée (style-visuel §2.1) : le nom de famille hérité de la face, tel que les
// fichiers statiques le déclarent (« Inter SemiBold », « Inter Tight SemiBold », « Inter 18pt
// SemiBold »). Sans elle, le demi-gras devient le gras de la famille.
function semiFace(name) {
  if (!name) return null;
  for (const cand of [name + ' SemiBold', name + ' 18pt SemiBold', name + ' DemiBold']) {
    if (fontInstalled(cand)) return cand;
  }
  return null;
}

// ---------------------------------------------------------------------------
// Thème
// ---------------------------------------------------------------------------
function resolveTheme(theme, mode, warnings, fontNotes) {
  theme = theme || {};
  const c = theme.couleurs || theme.colors || {};
  const T = {};
  T.nom = theme.nom || 'Theme';
  T.fondSombre = hex(c.fond_sombre, '111418', 'fond_sombre', warnings);
  T.fondClair = hex(c.fond_clair, 'FFFFFF', 'fond_clair', warnings);
  T.texte = hex(c.texte, '15181C', 'texte', warnings);
  T.accent = hex(c.accent, 'C2410C', 'accent', warnings);
  T.accent2 = hex(c.accent2, T.accent, 'accent2', warnings);
  // Dégradé du chiffre héros seulement si accent2 est donné et diffère de l'accent.
  T.gradient = !!c.accent2 && T.accent2 !== T.accent;
  const lightDefault = contrast(T.fondClair, T.fondSombre) >= 7 ? T.fondClair : 'FFFFFF';
  T.texteSombre = hex(c.texte_sur_sombre, lightDefault, 'texte_sur_sombre', warnings);
  // Texte secondaire : celui de l'utilisateur s'il se lit à 4,5:1, sinon 53 % du texte vers le fond.
  const secondary = (given, fg, bg, label) => {
    const auto = mix(fg, bg, STYLE.couleur.secondaire);
    const h = hex(given, null, label, warnings);
    if (!h) return auto;
    if (contrast(h, bg) >= 4.5) return h;
    warnings.push('Couleur « ' + label + ' » sous 4,5:1 sur son fond : gris calculé à la place (53 % du texte vers le fond).');
    return auto;
  };
  T.texteSec = secondary(c.texte_secondaire, T.texte, T.fondClair, 'texte_secondaire');
  T.texteSecSombre = secondary(c.texte_secondaire_sur_sombre, T.texteSombre, T.fondSombre, 'texte_secondaire_sur_sombre');
  T.pointAccent = theme.point_accent !== false;

  const pols = theme.polices || theme.fonts || {};
  const ft = fontSpec(pols.titre || pols.title, 'Arial');
  const fx = fontSpec(pols.texte || pols.body || pols.titre, 'Arial');
  // Le .pptx écrit toujours la police demandée : il sera projeté sur la machine de l'utilisateur, pas
  // sur celle qui le construit. Seul --polices repli écrit le repli. Si la police manque ici, l'aperçu
  // s'affiche dans le repli, et la mesure des textes prend la plus large des deux polices.
  const absentes = [];
  const replis = [];
  const pick = (f) => {
    if (mode === 'repli') return { ecrite: f.repli, ici: f.repli };
    if (f.nom === f.repli || fontInstalled(f.nom)) return { ecrite: f.nom, ici: f.nom };
    if (!absentes.includes(f.nom)) absentes.push(f.nom);
    if (!replis.includes(f.repli)) replis.push(f.repli);
    return { ecrite: f.nom, ici: f.repli };
  };
  const pt = pick(ft);
  const px = pick(fx);
  if (absentes.length) {
    const q = (l) => l.map((n) => '« ' + n + ' »').join(' et ');
    const s = absentes.length > 1;
    fontNotes.push('Police' + (s ? 's ' : ' ') + q(absentes) + ' absente' + (s ? 's' : '') + ' de cette machine : le .pptx ' + (s ? 'les écrit' : 'l\'écrit') + ' quand même, l\'aperçu s\'affiche en '
      + q(replis) + ' ; installez-' + (s ? 'les' : 'la') + ' sur la machine qui présente.');
  }
  T.fontTitre = pt.ecrite;
  T.fontTexte = px.ecrite;
  T.semiTitre = semiFace(T.fontTitre);
  T.semiTexte = semiFace(T.fontTexte);
  T.cleTitre = fontKey(T.fontTitre);
  T.cleTexte = fontKey(T.fontTexte);
  T.mesureTitre = uniq([T.cleTitre, fontKey(pt.ici)]);
  T.mesureTexte = uniq([T.cleTexte, fontKey(px.ici)]);
  T.iciTitre = fontKey(pt.ici);
  T.iciTexte = fontKey(px.ici);
  T.cssTitre = cssFamily(T.fontTitre, ft);
  T.cssTexte = cssFamily(T.fontTexte, fx);
  T.logo = theme.logo || null;
  return T;
}
function cssFamily(written, f) {
  const q = (n) => '\'' + String(n).replace(/'/g, '') + '\'';
  return uniq([written, f.nom, f.repli, 'Arial']).map(q).concat(['sans-serif']).join(', ');
}
// Graisse d'un paragraphe → face écrite dans le .pptx, gras ou non, et graisse CSS de l'aperçu.
function faceOf(T, role, wt) {
  const fam = role === 'titre' ? T.fontTitre : T.fontTexte;
  const semi = role === 'titre' ? T.semiTitre : T.semiTexte;
  if (wt === 'semi') return semi ? { face: semi, bold: false, css: 600 } : { face: fam, bold: true, css: 700 };
  if (wt === 'bold') return { face: fam, bold: true, css: 700 };
  return { face: fam, bold: false, css: 400 };
}
function palette(T, fond) {
  const C = STYLE.couleur;
  let P;
  if (fond === 'clair') {
    const bg = T.fondClair;
    P = {
      fond: 'clair', bg, fg: T.texte, fg2: T.texteSec, accent: T.accent, accent2: T.accent2,
      // Tuiles : blanches sur un fond gris clair ; gris très clair si le fond est déjà blanc.
      surface: contrast(bg, 'FFFFFF') > 1.04 ? mix(bg, 'FFFFFF', 0.8) : mix(bg, T.texte, 0.05)
    };
  } else if (fond === 'accent') {
    // Pleine page en couleur d'accent (slide finale) : l'accent du texte devient la couleur du texte.
    const bg = T.accent;
    const fg = readableOn(bg, T.fondSombre, T.texteSombre);
    P = { fond: 'accent', bg, fg, fg2: mix(fg, bg, 0.35), accent: fg, accent2: fg, surface: mix(bg, fg, 0.1) };
  } else {
    const bg = T.fondSombre;
    P = {
      fond: fond === 'image' ? 'image' : 'sombre', bg, fg: T.texteSombre, fg2: T.texteSecSombre,
      accent: T.accent, accent2: T.accent2, surface: mix(bg, T.texteSombre, C.tuileSombre)
    };
  }
  P.other = mix(P.fg, P.bg, C.autre);
  P.line = mix(P.fg, P.bg, 0.6);
  P.accentTxt = textAccent(P.bg, P.accent, P.fg);
  P.accent2Txt = textAccent(P.bg, P.accent2, P.fg);
  return P;
}

// ---------------------------------------------------------------------------
// Table des publics (même fichier que check_deck.py)
// ---------------------------------------------------------------------------
function loadPublics() {
  return readJSON(path.join(DATA_DIR, 'publics-exigences.json'), { defaut: 'defaut', publics: {}, prudence: {} });
}
function matchPublic(label, pubs) {
  const k = normKey(label);
  if (!k) return null;
  if (pubs[k]) return k;
  let best = null;
  let bestLen = 0;
  for (const key of Object.keys(pubs)) {
    for (const syn of [key].concat(pubs[key].synonymes || [])) {
      const sk = normKey(syn);
      if (!sk) continue;
      if (sk === k) return key;
      if (('-' + k + '-').includes('-' + sk + '-') && sk.length > bestLen) {
        best = key;
        bestLen = sk.length;
      }
    }
  }
  return best;
}
// meta.public : texte, ou { type, description } ; type peut être une liste (public mixte) :
// chaque réglage prend alors la valeur la plus prudente (ordre « prudence » du fichier).
function resolvePublic(meta, table) {
  const raw = meta && meta.public;
  let labels = [];
  if (raw && typeof raw === 'object' && !Array.isArray(raw)) labels = arr(raw.type || raw.types || raw.description);
  else labels = arr(raw);
  const pubs = table.publics || {};
  const keys = [];
  labels.forEach((l) => {
    const k = matchPublic(l, pubs);
    if (k && !keys.includes(k)) keys.push(k);
  });
  const reconnu = keys.length > 0;
  if (!keys.length) keys.push(table.defaut && pubs[table.defaut] ? table.defaut : Object.keys(pubs)[0] || 'defaut');
  const out = Object.assign({}, pubs[keys[0]] || {});
  const order = table.prudence || {};
  keys.slice(1).forEach((k) => {
    const row = pubs[k] || {};
    Object.keys(order).forEach((field) => {
      const ranks = order[field];
      if (ranks.indexOf(row[field]) > ranks.indexOf(out[field])) out[field] = row[field];
    });
  });
  return Object.assign(out, { cle: keys.join('+'), reconnu });
}

// ---------------------------------------------------------------------------
// Composition d'une slide : liste d'éléments positionnés, rendus ensuite en .pptx et en HTML.
// Chaque élément (ou paragraphe) porte un groupe d'apparition « g » (titre, lignes…) et,
// pour les listes, un rang « sub » : c'est ce que les animations ciblent.
// ---------------------------------------------------------------------------
class SlideCtx {
  constructor(deck, slide, index, env) {
    this.deck = deck;
    this.slide = slide;
    this.index = index;
    this.env = env;
    this.T = env.T;
    this.lang = env.lang;
    this.L = LABELS[env.lang] || LABELS.fr;
    this.t = slide.textes || {};
    const arch = String(slide.archetype || '').toUpperCase();
    this.arch = arch;
    this.renderArch = arch;
    let fond = String(slide.fond || '').toLowerCase();
    if (!['sombre', 'clair', 'image', 'accent'].includes(fond)) fond = 'sombre';
    // Slide finale d'un mot (A1) : pleine page en couleur d'accent, la seule du deck, si le texte
    // s'y lit à 4,5:1 (style-visuel, « Slide finale »).
    this.isFinal = index === env.finalIndex && arch === 'A1';
    if (this.isFinal && STYLE.finale.pleinePageAccent && (fond === 'sombre' || fond === 'clair')) {
      const PA = palette(this.T, 'accent');
      if (contrast(PA.bg, PA.fg) >= STYLE.finale.contraste) fond = 'accent';
    }
    this.fond = fond;
    this.P = palette(this.T, fond);
    this.items = [];
    this.warnings = [];
    this.proposals = [];   // [À VALIDER] affichés à l'écran
    this.fournirs = [];    // libellés [À FOURNIR] affichés à l'écran
    this.missing = [];     // images manquantes (placeholder dessiné)
    this.phs = [];         // placeholders dessinés (groupés dans le .pptx)
    this.names = {};
    this.curPh = null;
    this.accentWords = false;   // un mot *en accent* dans un texte de la slide
    this.accentUsed = false;    // l'accent est déjà pris (chiffre, barre, coches, ligne active) : surtitre en gris
    this.sourceText = slide.source ? String(slide.source) : '';
    const mode = env.sourceMode;
    // sous_le_chiffre : la source en clair sous le chiffre ; renvoi : un renvoi numéroté sous le
    // chiffre et la source complète au carton final ; carton_final : rien sur la slide.
    this.sourceOnSlide = !!this.sourceText && (mode === 'sous_le_chiffre' || mode === 'renvoi');
    this.sourceRef = 0;
    this.sourceDisplay = this.sourceText;
    if (this.sourceText && (mode === 'renvoi' || mode === 'carton_final')) {
      // Une même source garde le même numéro.
      const known = env.sourceIndex[this.sourceText];
      this.sourceNew = !known;
      if (!known) {
        env.sourceCounter += 1;
        env.sourceIndex[this.sourceText] = env.sourceCounter;
      }
      this.sourceRef = env.sourceIndex[this.sourceText];
      if (mode === 'renvoi') this.sourceDisplay = this.L.sourceRef(this.sourceRef);
    }
    this.sourceDone = false;
  }

  warn(msg) {
    if (!this.warnings.includes(msg)) this.warnings.push(msg);
  }

  // Ajoute un élément, avec un nom unique sur la slide (volet de sélection de PowerPoint).
  add(it) {
    const base = it.name || it.k;
    const n = (this.names[base] || 0) + 1;
    this.names[base] = n;
    it.name = n > 1 ? base + ' ' + n : base;
    if (this.curPh !== null && it.ph === undefined) {
      it.ph = this.curPh;
      const ph = this.phs[this.curPh];
      if (it.g === undefined) {
        it.g = ph.g;
        it.sub = ph.sub;
      }
    }
    this.items.push(it);
    return it;
  }

  // Champ de texte : valeur ou marqueur visible. Le champ principal d'une slide (titre, lignes,
  // chiffre, citation) qui manque prend le libellé tiré de la slide (hint), pas le nom du champ :
  // « titre », « lignes » ou « contenu » à l'écran sont des mots de l'atelier.
  f(field, label) {
    const v = this.t[field];
    if (v === undefined || v === null || String(v).trim() === '') {
      if (!this.hintUsed && ['titre', 'lignes', 'chiffre', 'citation'].includes(field)) {
        const h = this.hint();
        if (h) {
          this.hintUsed = true;
          label = h;
        }
      }
      return this.lang === 'en' ? '[TO PROVIDE: ' + label + ']' : '[À FOURNIR : ' + label + ']';
    }
    return String(v);
  }

  // Libellé court (3 mots au plus) de ce que la slide doit montrer : le début de textes.attendu,
  // sinon, pour une slide AF, son identifiant (« s13-garanties » → « Garanties »). Jamais un
  // libellé qui commence par un nombre (« 3 ou 4 garanties » se lirait comme un chiffre).
  hint() {
    const ok = (s) => {
      s = str(s).replace(/\s+/g, ' ').trim();
      const n = s.split(' ').filter(Boolean).length;
      return s && n <= 3 && !/^\d/.test(s) && !/[[\]*]/.test(s) ? s : '';
    };
    // « La slide de démonstration, à préparer… » → « démonstration » : ni article, ni « slide de ».
    const lead = /^(?:(?:la|le|les|un|une|des|the|an?)\s+|l['’]\s*)?(?:(?:slide|diapo|écran|ecran|screen)s?\s+(?:(?:de|du|des|of|for)\s+|d['’]\s*))?/i;
    let h = ok(str(this.t.attendu).split(/\s*[:;,.(]\s/)[0].replace(lead, ''));
    if (!h && this.arch === 'AF' && this.slide.id) {
      // Un identifiant qui n'est qu'un code (« af », « a9 », « s12 ») ne dit rien : on le laisse.
      const id = String(this.slide.id).replace(/^[a-z]?\d+[-_ ]*/i, '').replace(/[-_]+/g, ' ').trim();
      h = /[a-zà-ÿ]{3,}/i.test(id) && !/^(?:af|a\d+|slide|diapo)$/i.test(id) ? ok(id) : '';
      if (h) h = h.charAt(0).toUpperCase() + h.slice(1);
    }
    return h;
  }

  // Paragraphes d'un texte (un par ligne). o : rel (proportion du corps de la zone) ou pt (corps
  // fixe, ajusté à la police), gap (espace après, en pouces) ou gapK (en caps), wt (semi, regular,
  // bold), color, font (titre, texte), align, track, lh (interligne en caps), g, sub, subLines,
  // point (point final en accent), grad (dégradé du chiffre héros).
  para(text, o) {
    o = o || {};
    const lines = str(text).split(/\n/);
    return lines.map((line, li) => {
      const raw = stripValider(line, this.proposals);
      fournirLabels(raw).forEach((l) => { if (l && !this.fournirs.includes(l)) this.fournirs.push(l); });
      let runs = parseInline(typo(raw, this.lang), this.lang);
      if (runs.some((r) => r.accent)) this.accentWords = true;
      if (o.point && this.T.pointAccent) runs = accentPoint(runs);
      // Dégradé du chiffre héros : jamais sur un [À FOURNIR], qui reste estompé, d'une seule couleur.
      if (o.grad) runs = runs.map((r) => (r.marker ? r : Object.assign({}, r, { grad: true })));
      return {
        runs, rel: o.rel || 1, pt: o.pt ? scalePt(o.pt, o.font || 'texte') : undefined, gap: o.gap || 0, gapK: o.gapK,
        wt: o.wt, color: o.color, font: o.font, align: o.align, track: o.track, lh: o.lh, g: o.g, sub: o.subLines ? li : o.sub
      };
    });
  }

  // Corps le plus grand qui tient. o.stages : essais successifs { w, size, min, maxLines, maxTotal }
  // (largeur de mesure, corps, plancher, lignes permises par paragraphe et en tout).
  fitParas(item, o) {
    const stages = o.stages || [{ size: o.size, min: o.min, maxLines: o.maxLines, maxTotal: o.maxTotal }];
    for (const st of stages) {
      const size = scalePt(st.size !== undefined ? st.size : o.size || 18, item.font);
      const min = scalePt(st.min !== undefined ? st.min : (o.min !== undefined ? o.min : Math.min(st.size || o.size || 18, 12)), item.font);
      const probe = Object.assign({}, item, { w: st.w || item.w, size, min });
      const s = fitSize(probe, { maxLines: o.noWrap ? 1 : st.maxLines, maxTotal: st.maxTotal });
      if (s !== null) return s;
    }
    return null;
  }
  // Corps qui tient, sans rien ajouter à la slide (null si rien ne tient).
  fit(o) {
    const font = o.font || 'texte';
    const item = { paras: o.paras, w: o.w, h: o.h || 99, font, wt: o.wt || STYLE.graisse.texte, lh: o.lh || STYLE.interligne.texte, min: 0 };
    return this.fitParas(item, o);
  }

  // Ajoute une zone de texte ajustée à sa boîte. Corps (size, min, stages) en points de référence ;
  // fixed : corps final, sans ajustement.
  text(o) {
    const font = o.font || 'texte';
    const paras = o.paras || this.para(o.text === undefined ? '' : o.text, { color: o.color, wt: o.wt, font });
    const stages = o.stages || [{ size: o.size || 18, min: o.min, maxLines: o.maxLines, maxTotal: o.maxTotal }];
    const last = stages[stages.length - 1];
    const lastSize = last.size !== undefined ? last.size : (o.size || 18);
    const item = {
      k: 'text', x: o.x, y: o.y, w: o.w, h: o.h, paras,
      size: scalePt(stages[0].size !== undefined ? stages[0].size : (o.size || 18), font),
      min: scalePt(last.min !== undefined ? last.min : (o.min !== undefined ? o.min : Math.min(lastSize, 12)), font),
      wt: o.wt || STYLE.graisse.texte, align: o.align || 'left', valign: o.valign || 'middle', color: o.color || this.P.fg,
      font, lh: o.lh || STYLE.interligne.texte, name: o.name || 'Texte', g: o.g, sub: o.sub, noWrap: !!o.noWrap
    };
    if (o.fixed) item.fitted = o.fixed;
    else {
      item.fitted = this.fitParas(item, Object.assign({}, o, { stages }));
      if (item.fitted === null) {
        this.warn('Texte trop long pour sa place (' + item.name + ') : raccourcissez-le.');
        // En dernier recours, sous le plancher, pour ne pas sortir de la slide.
        item.fitted = fitSize(Object.assign({}, item, { size: item.min, min: Math.max(9, item.min * 0.6) }), { maxLines: o.noWrap ? 1 : undefined }) || item.min;
      }
    }
    if (o.balance && !o.noWrap) balanceParas(item);
    return this.add(item);
  }

  // Une ligne posée par sa ligne de base, au corps final o.size (sans ajustement).
  line(o) {
    const font = o.font || 'texte';
    const paras = o.paras || this.para(o.text, { color: o.color, wt: o.wt, font });
    const lh = o.lh || STYLE.interligne.titre;
    const cap = pol(font).cap * o.size / 72;
    const L = lh * cap;
    const h = L + 2 * PAD;
    return this.add({
      k: 'text', x: o.x, y: o.base - cap / 2 - h / 2, w: o.w, h, paras, size: o.size, min: o.size, fitted: o.size,
      wt: o.wt || STYLE.graisse.texte, align: o.align || 'center', valign: 'middle', color: o.color || this.P.fg,
      font, lh, name: o.name || 'Texte', g: o.g, sub: o.sub, noWrap: !!o.noWrap
    });
  }

  // Surtitre (style-visuel §3, « Surtitre ») : 0,42 × le corps du titre, entre 16 et 28 pt, sa ligne
  // de base à 0,45 cap au-dessus du haut des capitales du titre (geo : textGeom du titre). En accent
  // et demi-gras si l'accent de la slide est libre ; sinon en gris secondaire, regular.
  kicker(geo, o) {
    o = o || {};
    const s = this.t.surtitre;
    if (!s || !String(s).trim() || !geo) return null;
    const S = STYLE.surtitre;
    const size = o.size ? scalePt(o.size, 'texte')
      : clamp(Math.round(geo.firstSize * S.rapport * 2) / 2, scalePt(S.min, 'texte'), scalePt(S.max, 'texte'));
    const accent = o.accent !== undefined ? o.accent : !(this.accentUsed || this.accentWords);
    const color = o.color || (accent ? this.P.accentTxt : this.P.fg2);
    const base = geo.capTop - (o.ecart !== undefined ? o.ecart : S.ecart) * geo.capH;
    return this.line({ text: s, x: o.x, w: o.w, base, size, wt: o.wt || (accent ? 'semi' : 'regular'), color, align: o.align, font: 'texte', name: 'Surtitre', g: o.g || 'surtitre' });
  }

  rect(o) {
    return this.add(Object.assign({ k: 'rect' }, o));
  }
  round(o) {
    return this.add(Object.assign({ k: 'round' }, o));
  }
  ellipse(o) {
    return this.add(Object.assign({ k: 'ellipse' }, o));
  }
  check(o) {
    return this.add(Object.assign({ k: 'check' }, o));
  }

  imageSpec(spec) {
    if (!spec) return null;
    if (typeof spec === 'string') return { chemin: spec };
    return spec;
  }

  // Image dans une boîte ; si elle manque, placeholder dessiné aux couleurs du thème.
  // o : g, sub (groupe d'apparition), kind (forcer un type de placeholder), full (plein cadre),
  // alignX / alignY (left, right, top, bottom ; centre par défaut), radius (coins arrondis),
  // phBox(kind) (boîte du placeholder, si elle diffère de celle de l'image).
  // Retourne le rectangle dessiné de l'image, ou null quand un placeholder la remplace.
  image(box, spec, mode, name, o) {
    o = o || {};
    spec = this.imageSpec(spec);
    const attendu = spec ? str(spec.attendu || spec.libelle || spec.alt) : '';
    if (spec && spec.chemin) {
      const abs = path.resolve(this.env.deckDir, spec.chemin);
      const info = imageInfo(abs);
      if (info && info.w > 0 && info.h > 0) {
        const m = spec.ajustement === 'couvrir' ? 'cover' : spec.ajustement === 'contenir' ? 'contain' : (mode || 'contain');
        let r = { x: box.x, y: box.y, w: box.w, h: box.h };
        if (m === 'contain') {
          // Rectangle exact de l'image entière dans sa boîte : le même dans le .pptx et l'aperçu.
          const s = Math.min(box.w / info.w, box.h / info.h);
          const dw = info.w * s;
          const dh = info.h * s;
          r = {
            x: o.alignX === 'left' ? box.x : o.alignX === 'right' ? box.x + box.w - dw : box.x + (box.w - dw) / 2,
            y: o.alignY === 'top' ? box.y : o.alignY === 'bottom' ? box.y + box.h - dh : box.y + (box.h - dh) / 2,
            w: dw, h: dh
          };
        }
        this.add(Object.assign({ k: 'image', abs, info, mode: m === 'cover' ? 'cover' : 'fill', radius: o.radius || 0, name: name || 'Image', alt: attendu || path.basename(abs), g: o.g, sub: o.sub }, r));
        return r;
      }
      const reason = fs.existsSync(abs) ? 'illisible (PNG, JPEG ou GIF attendu)' : this.L.missingFile;
      this.warn('Image « ' + spec.chemin + ' » ' + reason + ' : placeholder dessiné à la place, à remplacer.');
    } else {
      this.warn('Image attendue' + (attendu ? ' (« ' + attendu + ' »)' : '') + ' : placeholder dessiné à la place.');
    }
    const kind = placeholderKind(this, spec, attendu, o);
    this.missing.push({ kind, attendu, chemin: spec && spec.chemin ? spec.chemin : '' });
    const pbox = (o.phBox && o.phBox(kind)) || box;
    drawPlaceholder(this, pbox, kind, attendu, Object.assign({ name }, o));
    return null;
  }

  sourceLine(o) {
    if (!this.sourceOnSlide || this.sourceDone) return;
    o = o || {};
    this.text({ x: o.x === undefined ? M : o.x, y: o.y === undefined ? 5.02 : o.y, w: o.w || W - 2 * M, h: o.h || 0.34, text: this.sourceDisplay, size: STYLE.source[0], min: STYLE.source[1], color: this.P.fg2, align: o.align || 'center', valign: 'top', font: 'texte', lh: STYLE.interligne.titre, name: 'Source', g: o.g });
    this.sourceDone = true;
  }
}

// Mesure ---------------------------------------------------------------------
function roleOf(it, p) {
  return (p && p.font) || it.font || 'texte';
}
function paraSize(it, p, s) {
  return p.pt ? p.pt : s * (p.rel || 1);
}
function paraWeight(it, p) {
  return (p && p.wt) || it.wt || STYLE.graisse.texte;
}
function heavy(wt) {
  return wt === 'semi' || wt === 'bold';
}
// Corps de référence (Inter) → corps qui donne la même hauteur de capitale dans la police du rôle.
function scalePt(pt, role) {
  if (pt === undefined || pt === null) return pt;
  return Math.round(pt * STYLE.capReference / pol(role).cap * 2) / 2;
}
// Suivi (em) selon la police et le corps (style-visuel §2.2).
function trackFor(role, ps) {
  if (ps < STYLE.suivi.sansSerrageSous) return 0;
  const s = pol(role).suivi;
  return Math.max(STYLE.suivi.max, ps >= 50 ? s[0] : ps >= 28 ? s[1] : s[2]);
}
function paraTrack(it, p, ps) {
  return p.track !== undefined ? p.track : trackFor(roleOf(it, p), ps);
}
function capIn(role, ps) {
  return pol(role).cap * ps / 72;
}
// Interligne exact (pouces) : k × hauteur de capitale.
function paraLineH(it, p, ps) {
  return (p.lh || it.lh || STYLE.interligne.texte) * capIn(roleOf(it, p), ps);
}
function paraGap(it, p, ps) {
  return p.gapK !== undefined ? p.gapK * capIn(roleOf(it, p), ps) : (p.gap || 0);
}
function paraInfo(it, p, ps, w) {
  const textOnly = p.runs.map((r) => r.t).join('');
  return wrapInfo(textOnly, ps, heavy(paraWeight(it, p)), w || it.w, paraTrack(it, p, ps), roleOf(it, p));
}
function layoutHeight(it, s, opts) {
  opts = opts || {};
  let hTot = 0;
  let ok = true;
  let total = 0;
  it.paras.forEach((p, i) => {
    const ps = paraSize(it, p, s);
    const info = paraInfo(it, p, ps);
    if (info.maxWord > it.w) ok = false;
    if (opts.maxLines && info.lines > opts.maxLines) ok = false;
    total += info.lines;
    hTot += info.lines * paraLineH(it, p, ps);
    if (i < it.paras.length - 1) hTot += paraGap(it, p, ps);
  });
  if (opts.maxTotal && total > opts.maxTotal) ok = false;
  return { h: hTot, ok, lines: total };
}
function fitSize(it, opts) {
  opts = opts || {};
  const floor = Math.max(it.min || 0, opts.floor || 0);
  for (let s = it.size; s >= floor - 1e-9; s -= (s > 48 ? 2 : 1)) {
    const r = layoutHeight(it, s, opts);
    if (r.ok && r.h <= it.h + 1e-6) return Math.round(s * 2) / 2;
  }
  return null;
}

// Géométrie verticale d'une zone de texte ajustée (pouces) : ses lignes, le haut des capitales de
// la première, la ligne de base de la dernière. Le centre d'une ligne est celui de ses capitales.
const PAD = 0.06;
function textGeom(it) {
  const lines = [];
  let y = 0;
  it.paras.forEach((p, i) => {
    const ps = paraSize(it, p, it.fitted);
    const L = paraLineH(it, p, ps);
    const cap = capIn(roleOf(it, p), ps);
    const n = it.noWrap ? 1 : paraInfo(it, p, ps).lines;
    for (let j = 0; j < n; j++) {
      lines.push({ c: y + L / 2, cap, L, ps });
      y += L;
    }
    if (i < it.paras.length - 1) y += paraGap(it, p, ps);
  });
  if (!lines.length) lines.push({ c: 0, cap: 0, L: 0, ps: it.fitted || 0 });
  const top = it.valign === 'top' ? it.y : it.valign === 'bottom' ? it.y + it.h - y : it.y + (it.h - y) / 2;
  lines.forEach((l) => { l.c += top; });
  const a = lines[0];
  const z = lines[lines.length - 1];
  return { top, h: y, n: lines.length, lines, capTop: a.c - a.cap / 2, capH: a.cap, firstSize: a.ps, base: z.c + z.cap / 2 };
}
// Zone de texte centrée sur cy (centre de la zone des capitales), resserrée sur son texte.
function centerBox(c, o) {
  const room = o.room !== undefined ? o.room : 2 * Math.max(0.6, Math.min(o.cy - 0.15, H - 0.15 - o.cy));
  const it = c.text(Object.assign({}, o, { y: o.cy - room / 2, h: room, valign: 'middle' }));
  const g = textGeom(it);
  it.h = g.h + 2 * PAD;
  it.y = o.cy - it.h / 2;
  return { it, g: textGeom(it) };
}
// Zone de texte posée par le haut des capitales de sa première ligne.
function topBox(c, o) {
  const it = c.text(Object.assign({}, o, { y: 0, h: o.room || H, valign: 'top' }));
  return placeTop(it, o.capTop);
}
function placeTop(it, capTop) {
  const g = textGeom(Object.assign({}, it, { y: 0, valign: 'top' }));
  it.valign = 'middle';
  it.h = g.h + 2 * PAD;
  it.y = capTop - g.capTop - PAD;
  return { it, g: textGeom(it) };
}
function shiftItems(items, dy) {
  items.forEach((it) => {
    if (it.k === 'check') it.cy += dy;
    else it.y += dy;
  });
}

// Équilibre les lignes d'un texte d'affichage : même nombre de lignes, largeur la plus
// régulière possible (évite un mot seul en dernière ligne). Les coupures deviennent des
// paragraphes, identiques dans le .pptx et dans l'aperçu.
function balanceParas(it) {
  const out = [];
  it.paras.forEach((p) => {
    const ps = paraSize(it, p, it.fitted);
    const hv = heavy(paraWeight(it, p));
    const tr = paraTrack(it, p, ps);
    const role = roleOf(it, p);
    const full = p.runs.map((r) => r.t).join('');
    const n = wrapInfo(full, ps, hv, it.w, tr, role).lines;
    if (n < 2 || n > 4) {
      out.push(p);
      return;
    }
    const lines = evenLines(full, ps, hv, it.w, n, tr, role);
    if (lines.length !== n) {
      out.push(p);
      return;
    }
    // Répartit les runs (accent, marqueurs) sur les lignes trouvées.
    const pieces = splitRuns(p.runs, lines);
    pieces.forEach((runs, i) => out.push(Object.assign({}, p, { runs, gap: i < pieces.length - 1 ? 0 : p.gap, gapK: i < pieces.length - 1 ? undefined : p.gapK })));
  });
  it.paras = out;
}
// Mots-outils qui ne finissent pas une ligne d'affichage (articles, prépositions, négation…).
const WEAK_END = /^(?:le|la|les|un|une|des|du|de|à|au|aux|en|et|ou|ne|n'|par|pour|sur|sous|sans|avec|dans|chez|vers|son|sa|ses|ce|cet|cette|ces|votre|vos|notre|nos|leur|leurs|mon|ma|mes|qui|que|dont|plus|a|an|the|of|to|in|on|at|for|with|and|or|by|from|your|our|my|its|their|is|are)$/i;
// Coupe en exactement n lignes, chacune <= w, les plus régulières possible
// (somme des écarts au carré à la longueur moyenne la plus petite).
function evenLines(s, pt, hv, w, n, track, role) {
  const words = s.split(/ +/).filter((x) => x.length);
  if (words.length < n || words.length > 40) return greedyLines(s, pt, hv, w, track, role);
  const ww = words.map((x) => textWidth(x, pt, hv, track, role));
  const sp = textWidth(' ', pt, hv, track, role);
  const lineW = (i, j) => {
    let t = 0;
    for (let k = i; k < j; k++) t += ww[k] + (k > i ? sp : 0);
    return t;
  };
  const total = lineW(0, words.length);
  const mean = (total - sp * (n - 1)) / n;
  const memo = new Map();
  const best = (i, left) => {
    const key = i + ':' + left;
    if (memo.has(key)) return memo.get(key);
    let res = null;
    if (left === 1) {
      const lw = lineW(i, words.length);
      res = lw <= w ? { cost: (lw - mean) * (lw - mean), cuts: [] } : null;
    } else {
      for (let j = i + 1; j <= words.length - (left - 1); j++) {
        const lw = lineW(i, j);
        if (lw > w) break;
        const rest = best(j, left - 1);
        if (!rest) continue;
        // Une ligne ne finit pas sur un mot-outil (« Une panne ne / prévient jamais. ») : la
        // coupure suit le sens, comme dans les keynotes, quitte à des lignes moins égales.
        const weak = WEAK_END.test(words[j - 1].replace(/[«»"“”]/g, '')) ? STYLE.coupure.motOutil * mean * mean : 0;
        const cost = (lw - mean) * (lw - mean) + weak + rest.cost;
        if (!res || cost < res.cost) res = { cost, cuts: [j].concat(rest.cuts) };
      }
    }
    memo.set(key, res);
    return res;
  };
  const r = best(0, n);
  if (!r) return greedyLines(s, pt, hv, w, track, role);
  const out = [];
  let start = 0;
  r.cuts.concat([words.length]).forEach((end) => {
    out.push(words.slice(start, end).join(' '));
    start = end;
  });
  return out;
}
function greedyLines(s, pt, hv, w, track, role) {
  const words = s.split(/ +/).filter((x) => x.length);
  const sp = textWidth(' ', pt, hv, track, role);
  const lines = [];
  let cur = '';
  let curW = 0;
  for (const wd of words) {
    const ww = textWidth(wd, pt, hv, track, role);
    if (!cur) {
      cur = wd;
      curW = ww;
    } else if (curW + sp + ww <= w) {
      cur += ' ' + wd;
      curW += sp + ww;
    } else {
      lines.push(cur);
      cur = wd;
      curW = ww;
    }
  }
  if (cur) lines.push(cur);
  return lines;
}
function splitRuns(runs, lines) {
  // Coupe la suite de runs aux frontières de lignes (les espaces de coupure sont retirés).
  const flat = [];
  runs.forEach((r) => {
    for (const ch of r.t) flat.push({ ch, r });
  });
  const text = flat.map((x) => x.ch).join('').replace(/ +/g, ' ');
  if (text.trim() !== lines.join(' ')) return [runs];
  const out = [];
  let idx = 0;
  const chars = flat.filter((x, i) => !(x.ch === ' ' && flat[i - 1] && flat[i - 1].ch === ' '));
  while (idx < chars.length && chars[idx].ch === ' ') idx += 1;
  lines.forEach((line) => {
    const seg = chars.slice(idx, idx + line.length);
    idx += line.length;
    while (idx < chars.length && chars[idx].ch === ' ') idx += 1;
    const lr = [];
    seg.forEach((x) => {
      const last = lr[lr.length - 1];
      if (last && last.src === x.r) last.t += x.ch;
      else lr.push({ t: x.ch, src: x.r });
    });
    out.push(lr.map((r) => Object.assign({}, r.src, { t: r.t })));
  });
  return out;
}

// ---------------------------------------------------------------------------
// Placeholders dessinés (style-visuel §5.2) : une image qui manque devient une interface
// schématique, un appareil générique, un objet, une photo ou une carte, aux couleurs de la slide.
// Aucun chiffre ni texte inventé, aucun trait, aucune consigne à l'écran : des aplats, des barres
// grises et une seule ligne en accent. Ce que l'image doit montrer part dans les notes.
// ---------------------------------------------------------------------------
const KIND_RX = {
  appareil: /\b(mobile|telephones?|smartphones?|portables?|android|phones?|tablettes?|tablets?)\b/,
  carte: /\b(cartes? (?:de|du|des)|plans? de (?:la )?ville|zones? desservies?|maps?|territoires?)\b/,
  interface: /\b(captures?|ecrans?|interfaces?|tableaux? de bord|dashboards?|pages?|applis?|applications?|logiciels?|modules?|demos?|factures?|listes?|screens?|screenshots?|ui|web|sites?|espaces? clients?|portails?|suivis?)\b/,
  photo: /\b(photos?|portraits?|equipes?|personnes?|scenes?|decors?|ambiances?|paysages?|people|team|crowd|lifestyle)\b/,
  objet: /\b(produits?|boitiers?|objets?|capteurs?|machines?|camions?|vehicules?|emballages?|packagings?|bouteilles?|hardware|products?|box|coffrets?)\b/
};
const VARIANT_RX = [
  ['tableau', /\b(tableaux? de bord|dashboards?|statistiques?|kpi|indicateurs?|graphiques?|charts?|analytics|rapports?|reporting)\b/],
  ['liste', /\b(listes?|suivis?|statuts?|historiques?|lists?|tables?|inbox|commandes|toutes les)\b/],
  ['document', /\b(factures?|documents?|details?|fiches?|formulaires?|devis|bons? de|contrats?|invoices?|forms?|records?|e-mails?|mails?)\b/]
];
function normText(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
    .replace(/[’`]/g, '\'');
}
function placeholderKind(c, spec, attendu, o) {
  const forced = normText(spec && (spec.rendu || spec.placeholder || spec.type));
  const map = { interface: 'interface', capture: 'interface', ecran: 'interface', appareil: 'appareil', telephone: 'appareil', mobile: 'appareil', objet: 'objet', produit: 'objet', photo: 'photo', carte: 'carte', map: 'carte' };
  if (map[forced]) return map[forced];
  if (o.kind) return o.kind;
  const t = normText(attendu + ' ' + (spec && spec.chemin ? path.basename(spec.chemin).replace(/[-_.]/g, ' ') : ''));
  if (KIND_RX.appareil.test(t)) return 'appareil';
  if (KIND_RX.carte.test(t)) return 'carte';
  if (KIND_RX.interface.test(t)) return 'interface';
  if (KIND_RX.photo.test(t)) return 'photo';
  if (KIND_RX.objet.test(t)) return 'objet';
  if (['A13', 'A19', 'A20'].includes(c.renderArch)) return 'objet';
  return 'interface';
}
// Couleurs d'un schéma. L'écran se lit comme une vraie capture en mode clair, quel que soit le fond
// de la slide : sur fond sombre, une carte claire (fond_clair, ou blanc) ; sur fond clair, une
// carte blanche, ou grise à 5 % si le fond est déjà blanc, ses aplats intérieurs devenant blancs.
// Une carte sombre à 8 % (itération 2, première version) se lisait comme un écran de chargement.
function schemaColors(c) {
  const P = c.P;
  const T = c.T;
  const SC = STYLE.schema;
  const dark = lum(P.bg) < 0.35;
  const ink = lum(T.texte) < 0.35 ? T.texte : '15181C';
  let card;
  if (dark) card = contrast(T.fondClair, P.bg) >= 7 ? T.fondClair : 'FFFFFF';
  else card = contrast(P.bg, 'FFFFFF') > 1.04 ? 'FFFFFF' : mix(P.bg, ink, SC.carteGrise);
  const white = contrast(card, 'FFFFFF') < 1.04;
  const C = {
    dark, card, shadow: !dark,
    sub: white ? mix(card, ink, 0.045) : 'FFFFFF',   // aplats intérieurs (lignes, cartes, champs), sans trait
    bar: mix(card, ink, SC.barres[0]),
    strong: mix(card, ink, SC.barres[1]),
    soft: mix(card, ink, 0.1),
    accent: P.accent,
    bezel: dark ? mix(P.bg, P.fg, 0.2) : mix(P.bg, ink, 0.85)
  };
  // Un accent clair (jaune sur blanc : 1,1:1) disparaît en barre fine : il devient alors une
  // pastille pleine, avec un trait sombre dedans, et la ligne mise en avant se teinte d'accent.
  C.accentFaint = contrast(card, P.accent) < SC.contrasteBarre;
  C.onAccent = mix(onAccent(P.accent, ink), P.accent, 0.3);
  C.tint = mix(C.sub === 'FFFFFF' ? 'FFFFFF' : C.sub, P.accent, 0.22);
  return C;
}
// La ligne mise en avant d'un schéma : une barre d'accent, ou une pastille si l'accent est trop clair.
function hiBar(c, C, x, y, w, h) {
  if (!C.accentFaint) {
    pill(c, x, y, w, h, C.accent, 'Schéma ligne en avant');
    return;
  }
  const ph = h * 2.6;
  pill(c, x - ph * 0.35, y + h / 2 - ph / 2, w + ph * 0.7, ph, C.accent, 'Schéma ligne en avant');
  pill(c, x, y, w * 0.72, h, C.onAccent, 'Schéma ligne en avant trait');
}
function drawPlaceholder(c, box, kind, attendu, o) {
  const L = c.L;
  const cap = L.cap[kind] || L.cap.interface;
  const short = shortLabel(str(stripValider(attendu)).replace(new RegExp(FOURNIR_SRC, 'gi'), (m) => markerInner(m)), 6);
  // Nom du groupe dans le volet de sélection du .pptx : il dit quoi remplacer, sans rien afficher.
  const label = cap + (short ? ' · ' + short : '');
  const ph = { id: c.phs.length, name: label, kind, g: o.g, sub: o.sub };
  c.phs.push(ph);
  c.curPh = ph.id;
  try {
    const C = schemaColors(c);
    const rnd = rng(31 * (c.index + 1) + Math.round(box.x * 10) + Math.round(box.y * 7));
    if (kind === 'photo') drawPhoto(c, o.full ? { x: 0, y: 0, w: W, h: H } : box, C, !!o.full);
    else if (kind === 'objet') drawObject(c, box, C);
    else if (kind === 'carte') drawMap(c, box, C, rnd);
    else if (kind === 'appareil') drawPhone(c, box, C);
    else if (kind === 'icone') drawIcon(c, box);
    else if (kind === 'logo') drawLogo(c, box, C, attendu);
    // Tuile d'image (mur, bento) : un aplat voisin des tuiles, pas un écran clair au milieu du mur.
    else if (kind === 'tuile') c.round({ x: box.x, y: box.y, w: box.w, h: box.h, fill: C.dark ? mix(c.P.bg, c.P.fg, 0.16) : C.sub, radius: Math.min(STYLE.A14.rayon, box.w / 4, box.h / 4), name: 'Schéma' });
    else drawWindow(c, box, C, rnd, interfaceVariant(attendu, c.index));
  } finally {
    c.curPh = null;
  }
}
function interfaceVariant(attendu, index) {
  const t = normText(attendu);
  for (const [v, rx] of VARIANT_RX) if (rx.test(t)) return v;
  return ['liste', 'document', 'tableau'][index % 3];
}
function pill(c, x, y, w, h, fill, name) {
  if (w <= 0 || h <= 0) return;
  c.round({ x, y, w, h, fill, radius: h / 2, name: name || 'Schéma' });
}
// Fenêtre d'application : carte sans trait, un mot-titre en barre forte, une colonne de navigation
// si la place le permet, puis le contenu (liste, tableau de bord ou document).
function drawWindow(c, b, C, rnd, variant) {
  const SC = STYLE.schema;
  const bh = SC.barre;
  c.round({ x: b.x, y: b.y, w: b.w, h: b.h, fill: C.card, radius: Math.min(SC.rayon, b.w / 6, b.h / 6), shadow: C.shadow, name: 'Schéma fenêtre' });
  const pad = clamp(Math.min(b.w, b.h) * 0.07, 0.1, 0.3);
  if (b.h < 1.3 || b.w < 1.8) {
    // Petite carte : un mot-titre, trois barres, une en accent.
    pill(c, b.x + pad, b.y + pad, (b.w - 2 * pad) * 0.5, bh * 1.2, C.strong);
    for (let i = 0; i < 3; i++) {
      const y = b.y + pad + (i + 1.6) * SC.pas;
      if (y + bh > b.y + b.h - pad * 0.6) break;
      if (i === 1) hiBar(c, C, b.x + pad, y, (b.w - 2 * pad) * SC.longueurs[i], bh);
      else pill(c, b.x + pad, y, (b.w - 2 * pad) * SC.longueurs[i], bh, C.bar);
    }
    return;
  }
  pill(c, b.x + pad, b.y + pad, b.w * 0.18, bh * 1.3, C.strong);
  pill(c, b.x + b.w - pad - b.w * 0.14, b.y + pad - bh * 0.35, b.w * 0.14, bh * 2, C.sub);
  const y0 = b.y + pad + bh * 1.3 + pad;
  const bottom = b.y + b.h - pad;
  let cx = b.x + pad;
  let cw = b.w - 2 * pad;
  if (b.w >= 4.2) {
    const sw = b.w * 0.2;
    const lens = [0.85, 0.6, 0.75, 0.5, 0.7, 0.55, 0.65];
    for (let i = 0; i < 7 && y0 + i * 0.3 + bh <= bottom; i++) pill(c, b.x + pad, y0 + i * 0.3, (sw - pad * 1.5) * lens[i], bh, C.bar);
    cx = b.x + sw + pad * 0.5;
    cw = b.w - sw - pad * 1.5;
  }
  const area = { x: cx, y: y0, w: cw, h: bottom - y0 };
  if (area.h < 0.4) return;
  if (variant === 'tableau') drawDashboard(c, area, C, rnd);
  else if (variant === 'document') drawDocument(c, area, C, rnd);
  else drawList(c, area, C);
}
function drawList(c, a, C) {
  const SC = STYLE.schema;
  const rh = clamp(a.h / 6, 0.4, 0.6);
  const rows = Math.max(2, Math.floor((a.h + 0.07) / rh));
  const hi = Math.min(rows - 1, 1);
  const tw = a.w * 0.5;
  for (let i = 0; i < rows; i++) {
    const y = a.y + i * rh;
    const h = rh - 0.07;
    c.round({ x: a.x, y, w: a.w, h, fill: i === hi ? C.tint : C.sub, radius: Math.min(0.08, h / 3), name: i === hi ? 'Schéma ligne en avant' : 'Schéma ligne' });
    const d = Math.min(0.22, h * 0.5);
    c.ellipse({ x: a.x + 0.12, y: y + (h - d) / 2, w: d, h: d, fill: C.bar, name: 'Schéma' });
    const tx = a.x + 0.12 + d + 0.14;
    const ty = y + h / 2 - (SC.pas + SC.barre) / 2;
    pill(c, tx, ty, tw * SC.longueurs[i % 3], SC.barre, C.strong);
    pill(c, tx, ty + SC.pas, tw * SC.longueurs[(i + 1) % 3] * 0.6, SC.barre, C.bar);
    // Colonne d'état : une pastille par ligne, en accent sur la ligne mise en avant.
    const sw = a.w * 0.13;
    const sh = Math.min(0.2, h * 0.42);
    pill(c, a.x + a.w - 0.12 - sw, y + h / 2 - sh / 2, sw, sh, i === hi ? C.accent : C.soft, i === hi ? 'Schéma état en avant' : 'Schéma');
  }
}
function drawDashboard(c, a, C, rnd) {
  const SC = STYLE.schema;
  const g = 0.1;
  const kh = clamp(a.h * 0.3, 0.45, 0.8);
  const kw = (a.w - 2 * g) / 3;
  for (let k = 0; k < 3; k++) {
    const x = a.x + k * (kw + g);
    c.round({ x, y: a.y, w: kw, h: kh, fill: C.sub, radius: 0.08, name: 'Schéma carte' });
    pill(c, x + 0.12, a.y + 0.14, kw * 0.4, SC.barre, C.bar);
    pill(c, x + 0.12, a.y + kh - 0.14 - SC.barre * 2, kw * (0.5 + rnd() * 0.25), SC.barre * 2, C.strong);
  }
  const cy = a.y + kh + g;
  const ch = a.h - kh - g;
  if (ch < 0.5) return;
  c.round({ x: a.x, y: cy, w: a.w, h: ch, fill: C.sub, radius: 0.08, name: 'Schéma graphique' });
  pill(c, a.x + 0.14, cy + 0.14, a.w * 0.2, SC.barre, C.bar);
  const n = 9;
  const pad = 0.2;
  const step = (a.w - 2 * pad) / n;
  const bw = step * 0.5;
  const base = cy + ch - 0.16;
  const topY = cy + 0.4;
  for (let i = 0; i < n; i++) {
    const bh = Math.max(0.06, (base - topY) * (0.3 + rnd() * 0.7));
    c.round({ x: a.x + pad + i * step + bw * 0.5, y: base - bh, w: bw, h: bh, fill: i === 5 ? C.accent : C.bar, radius: Math.min(bw / 2, 0.05), name: i === 5 ? 'Schéma barre en avant' : 'Schéma' });
  }
}
function drawDocument(c, a, C, rnd) {
  const SC = STYLE.schema;
  const dw = a.w * 0.62;
  c.round({ x: a.x, y: a.y, w: dw, h: a.h, fill: C.sub, radius: 0.08, name: 'Schéma document' });
  const pad = 0.18;
  const iw = dw - 2 * pad;
  pill(c, a.x + pad, a.y + pad, iw * 0.4, SC.barre * 1.4, C.strong);
  for (let i = 0; i < 2; i++) pill(c, a.x + pad, a.y + pad + 0.3 + i * SC.pas, iw * 0.55 * SC.longueurs[i + 1], SC.barre, C.bar);
  let y = a.y + pad + 0.3 + 2 * SC.pas + 0.16;
  const stepY = SC.pas * 1.8;
  const rows = Math.max(2, Math.floor((a.y + a.h - pad - 0.32 - y) / stepY));
  const hi = Math.min(rows - 1, 1 + Math.floor(rnd() * 2));
  for (let i = 0; i < rows; i++) {
    if (i === hi) hiBar(c, C, a.x + pad, y, iw * 0.55 * SC.longueurs[i % 3], SC.barre);
    else pill(c, a.x + pad, y, iw * 0.55 * SC.longueurs[i % 3], SC.barre, C.bar);
    pill(c, a.x + pad + iw * 0.78, y, iw * 0.22, SC.barre, C.bar);
    y += stepY;
  }
  pill(c, a.x + pad + iw * 0.6, a.y + a.h - pad - SC.barre * 1.4, iw * 0.4, SC.barre * 1.4, C.strong);
  // Panneau latéral : champs en aplat, un bouton.
  const px = a.x + dw + 0.14;
  const pw = a.w - dw - 0.14;
  if (pw < 0.5) return;
  for (let i = 0; i < 3; i++) {
    const fy = a.y + i * 0.62;
    if (fy + 0.5 > a.y + a.h - 0.45) break;
    pill(c, px, fy + 0.04, pw * 0.45, SC.barre, C.bar);
    c.round({ x: px, y: fy + 0.2, w: pw, h: 0.28, fill: C.sub, radius: 0.06, name: 'Schéma champ' });
  }
  c.round({ x: px, y: a.y + a.h - 0.3, w: pw, h: 0.3, fill: C.strong, radius: 0.15, name: 'Schéma bouton' });
}
// Appareil générique : coins à 13 % de la largeur, bord de 3 %, ni encoche ni bouton.
function drawPhone(c, b, C) {
  const SC = STYLE.schema;
  const ph = b.h;
  const pw = Math.min(b.w * 0.92, ph * 0.49);
  const x0 = b.x + (b.w - pw) / 2;
  c.round({ x: x0, y: b.y, w: pw, h: ph, fill: C.bezel, radius: pw * SC.appareil.coin, shadow: C.shadow, name: 'Schéma appareil' });
  const e = pw * SC.appareil.bord;
  const sx = x0 + e;
  const sy = b.y + e;
  const sw = pw - 2 * e;
  const sh = ph - 2 * e;
  c.round({ x: sx, y: sy, w: sw, h: sh, fill: C.card, radius: pw * (SC.appareil.coin - SC.appareil.bord), name: 'Schéma écran' });
  const p = sw * 0.09;
  const ix = sx + p;
  const iw = sw - 2 * p;
  pill(c, ix, sy + sh * 0.09, iw * 0.55, SC.barre * 1.4, C.strong);
  pill(c, ix, sy + sh * 0.09 + SC.pas * 1.2, iw * 0.35, SC.barre, C.bar);
  const ct = sy + sh * 0.2;
  const cardH = sh * 0.12;
  const gap = sh * 0.025;
  for (let i = 0; i < 4; i++) {
    const y = ct + i * (cardH + gap);
    c.round({ x: ix, y, w: iw, h: cardH, fill: i === 1 ? C.tint : C.sub, radius: Math.min(0.1, cardH * 0.25), name: 'Schéma carte' });
    const d = cardH * 0.42;
    c.ellipse({ x: ix + cardH * 0.25, y: y + (cardH - d) / 2, w: d, h: d, fill: C.bar, name: 'Schéma' });
    const tx = ix + cardH * 0.25 + d + cardH * 0.22;
    const tw = ix + iw - tx - cardH * 0.25;
    const ty = y + cardH / 2 - (SC.pas * 0.8 + SC.barre) / 2;
    if (i === 1) hiBar(c, C, tx, ty, tw * SC.longueurs[i % 3], SC.barre);
    else pill(c, tx, ty, tw * SC.longueurs[i % 3], SC.barre, C.strong);
    pill(c, tx, ty + SC.pas * 0.8, tw * 0.5, SC.barre, C.bar);
  }
  pill(c, ix, sy + sh * 0.9, iw, Math.max(SC.barre * 2.6, sh * 0.055), C.strong, 'Schéma bouton');
}
// Objet : une silhouette simple et un voyant d'accent ; ombre de contact sur fond clair seulement.
function drawObject(c, b, C) {
  const s = Math.min(b.w, b.h) * 0.7;
  const cx = b.x + b.w / 2;
  const top = b.y + (b.h - s) / 2;
  if (!C.dark) c.ellipse({ x: cx - s * 0.5, y: top + s * 0.96, w: s, h: s * 0.1, fill: c.P.fg, transparency: 90, name: 'Schéma ombre' });
  c.round({ x: cx - s * 0.5, y: top, w: s, h: s, fill: mix(c.P.bg, c.P.fg, C.dark ? 0.16 : 0.1), radius: s * 0.22, name: 'Schéma objet' });
  c.round({ x: cx - s * 0.34, y: top + s * 0.2, w: s * 0.68, h: s * 0.42, fill: mix(c.P.bg, c.P.fg, C.dark ? 0.08 : 0.18), radius: s * 0.08, name: 'Schéma face' });
  const d = s * 0.07;
  c.ellipse({ x: cx - d / 2, y: top + s * 0.76, w: d, h: d, fill: C.accent, name: 'Schéma voyant' });
}
// Photo : un fond un peu plus clair que la slide et deux halos neutres. Plus de reflet d'accent :
// l'accent translucide sur un fond sombre virait au kaki.
function drawPhoto(c, b, C, full) {
  const ink = c.P.fg;
  const base = mix(c.P.bg, ink, C.dark ? 0.07 : 0.06);
  if (full) {
    c.rect({ x: b.x, y: b.y, w: b.w, h: b.h, fill: base, name: 'Schéma photo' });
    c.ellipse({ x: b.x + b.w * 0.45, y: b.y - b.h * 0.35, w: b.w * 0.75, h: b.h * 1.25, fill: ink, transparency: 93, name: 'Schéma lumière' });
    c.ellipse({ x: b.x - b.w * 0.15, y: b.y + b.h * 0.45, w: b.w * 0.7, h: b.h * 0.95, fill: ink, transparency: 95, name: 'Schéma lumière' });
  } else {
    c.round({ x: b.x, y: b.y, w: b.w, h: b.h, fill: base, radius: Math.min(STYLE.schema.rayon, b.w / 6, b.h / 6), name: 'Schéma photo' });
    c.ellipse({ x: b.x + b.w * 0.36, y: b.y + b.h * 0.1, w: b.w * 0.56, h: b.h * 0.76, fill: ink, transparency: 91, name: 'Schéma lumière' });
    c.ellipse({ x: b.x + b.w * 0.08, y: b.y + b.h * 0.42, w: b.w * 0.4, h: b.h * 0.48, fill: ink, transparency: 94, name: 'Schéma lumière' });
  }
}
function drawMap(c, b, C, rnd) {
  c.round({ x: b.x, y: b.y, w: b.w, h: b.h, fill: C.card, radius: Math.min(STYLE.schema.rayon, b.w / 6, b.h / 6), shadow: C.shadow, name: 'Schéma carte' });
  const t = 0.022;
  for (let i = 1; i < 6; i++) c.rect({ x: b.x + 0.1, y: b.y + b.h * i / 6 + (rnd() - 0.5) * 0.1, w: b.w - 0.2, h: t, fill: C.soft, name: 'Schéma rue' });
  for (let i = 1; i < 8; i++) c.rect({ x: b.x + b.w * i / 8 + (rnd() - 0.5) * 0.15, y: b.y + 0.1, w: t, h: b.h - 0.2, fill: C.soft, name: 'Schéma rue' });
  c.rect({ x: b.x + 0.1, y: b.y + b.h * 0.55, w: b.w - 0.2, h: t * 2, fill: C.bar, name: 'Schéma axe' });
  c.ellipse({ x: b.x + b.w * 0.3, y: b.y + b.h * 0.2, w: b.w * 0.42, h: b.h * 0.58, fill: c.P.accent, transparency: 75, name: 'Schéma zone' });
  const d = Math.min(b.w, b.h) * 0.07;
  c.ellipse({ x: b.x + b.w * 0.51 - d / 2, y: b.y + b.h * 0.49 - d / 2, w: d, h: d, fill: c.P.accent, name: 'Schéma repère' });
}
function drawIcon(c, b) {
  const s = Math.min(b.w, b.h);
  const x = b.x + (b.w - s) / 2;
  const y = b.y + (b.h - s) / 2;
  c.round({ x, y, w: s, h: s, fill: c.P.accent, radius: s * 0.24, name: 'Schéma icône' });
  c.ellipse({ x: x + s * 0.3, y: y + s * 0.3, w: s * 0.4, h: s * 0.4, line: onAccent(c.P.accent, c.T.fondSombre), lineW: Math.max(2, s * 6), name: 'Schéma' });
}
// Logo à fournir : son nom (ou son marqueur) dans le style du texte, sans cadre.
function drawLogo(c, b, C, nom) {
  if (!str(nom).trim()) {
    pill(c, b.x + b.w * 0.2, b.y + b.h / 2 - 0.06, b.w * 0.6, 0.12, C.bar);
    return;
  }
  c.text({ x: b.x, y: b.y, w: b.w, h: b.h, text: nom, size: STYLE.A16.nom[0], min: STYLE.A16.nom[1], maxTotal: 2, wt: 'semi', color: c.P.fg2, align: 'center', font: 'texte', lh: STYLE.interligne.titre, name: 'Logo nom' });
}

// ---------------------------------------------------------------------------
// Archétypes : les fiches de style-visuel.md §3. Chaque bloc de texte est posé par le centre de sa
// zone des capitales (centerBox) ou par le haut de ses capitales (topBox, placeTop) ; les mêmes
// positions servent au .pptx et à l'aperçu.
// ---------------------------------------------------------------------------
const R = {};
const K = STYLE.interligne;
const FULL_W = W - 2 * M;
// Texte tel qu'il s'affiche (marqueurs retirés), pour estimer une largeur.
function shownText(c, s) {
  return parseInline(typo(stripValider(s), c.lang), c.lang).map((r) => r.t).join('');
}
// Essais d'un mot ou d'un nom très gros (A1) : une ligne dans 6,4 in, puis une ligne large, puis deux lignes.
function bigStages(max, multi) {
  const A = STYLE.A1;
  if (multi) return [{ w: A.w, size: Math.min(max, A.deuxLignes[0]), min: A.deuxLignes[1], maxLines: 1 }];
  return [
    { w: A.uneLigne, size: max, min: A.corps[1], maxLines: 1 },
    { w: A.w, size: A.large[0], min: A.large[1], maxLines: 1 },
    { w: A.w, size: A.deuxLignes[0], min: A.deuxLignes[1], maxTotal: 2 }
  ];
}

R.A1 = (c) => {
  const A = STYLE.A1;
  const S = c.isFinal ? STYLE.finale : A;
  const paras = c.para(c.f('titre', 'titre'), { wt: 'semi', font: 'titre', g: 'titre', point: true });
  const t = centerBox(c, { x: A.x, w: A.w, cy: S.cy, paras, stages: bigStages(S.corps[0], paras.length > 1), wt: 'semi', align: 'center', font: 'titre', lh: K.titre, balance: true, name: 'Titre' });
  c.kicker(t.g, { x: A.x, w: A.w, align: 'center' });
  c.sourceLine();
};

R.A2 = (c) => {
  const S = STYLE.A2;
  const spec = c.slide.image || (c.T.logo ? { chemin: c.T.logo, attendu: 'logo' } : { attendu: c.lang === 'en' ? 'icon or logo' : 'icône ou logo' });
  const paras = c.para(c.f('titre', 'nom'), { wt: 'semi', font: 'titre', g: 'titre' });
  const t = topBox(c, { x: M, w: FULL_W, capTop: 0, paras, size: S.nom[0], min: S.nom[1], maxTotal: 2, wt: 'semi', align: 'center', font: 'titre', lh: K.titre, balance: true, name: 'Nom' });
  const top = S.cy - (S.icone + S.ecart + (t.g.base - t.g.capTop)) / 2;
  const g = placeTop(t.it, top + S.icone + S.ecart).g;
  c.image({ x: W / 2 - S.icone / 2, y: top, w: S.icone, h: S.icone }, spec, 'contain', 'Icône', { g: 'visuel', kind: 'icone' });
  c.kicker({ capTop: top, capH: g.capH, firstSize: g.firstSize }, { x: M, w: FULL_W, align: 'center', ecart: 0.9 });
  c.sourceLine();
};

// Lignes d'une liste : la ligne active (actif, de 1 à n) en activeColor, les autres en idleColor.
function listParas(c, lines, o) {
  const actif = Number(c.t.actif) || 0;
  const out = [];
  lines.forEach((l, i) => {
    let color = o.color;
    if (actif) color = i + 1 === actif ? o.activeColor : o.idleColor;
    c.para(l, { color, wt: 'semi', font: o.font, gapK: o.gapK, g: 'lignes', sub: i }).forEach((p) => out.push(p));
  });
  return out;
}

R.A3 = (c) => {
  const S = STYLE.A3;
  const lines = arr(c.t.lignes);
  if (!lines.length) lines.push(c.f('lignes', 'lignes'));
  c.accentUsed = !!(Number(c.t.actif) || 0);
  const paras = listParas(c, lines, { color: c.P.fg, activeColor: c.P.accentTxt, idleColor: c.P.fg2, font: 'titre' });
  const size = S.corps[clamp(lines.length, 1, 8)];
  const t = centerBox(c, { x: S.x, w: S.w, cy: S.cy, paras, size, min: Math.min(size, S.plancher), maxLines: 1, wt: 'semi', align: 'center', font: 'titre', lh: K.sommaire, name: 'Sommaire' });
  c.kicker(t.g, { x: S.x, w: S.w, align: 'center' });
  c.sourceLine();
};

R.A4 = (c) => {
  const S = STYLE.A4;
  // Avec des retours à la ligne voulus (\n), chaque ligne reste entière.
  const paras = c.para(c.f('titre', 'phrase'), { wt: 'semi', font: 'titre', g: 'titre', subLines: true, point: true });
  const stages = [{ size: S.corps[0], min: S.corps[1], maxLines: paras.length > 1 ? 1 : 2, maxTotal: 2 }];
  const t = centerBox(c, { x: S.x, w: S.w, cy: S.cy, paras, stages, wt: 'semi', align: 'center', font: 'titre', lh: K.titre, balance: true, name: 'Phrase' });
  c.kicker(t.g, { x: S.x, w: S.w, align: 'center' });
  c.sourceLine();
};

// A5 : chiffre héros centré sur le milieu, libellé en bas (mise en page a) ; si le libellé tient
// sur deux lignes, le bloc entier remonte et le libellé suit le chiffre (mise en page b).
R.A5 = (c) => {
  const S = STYLE.A5;
  c.accentUsed = true;
  const grad = c.T.gradient && c.P.accent === c.T.accent && contrast(c.P.bg, c.T.accent) >= 2.6 && contrast(c.P.bg, c.T.accent2) >= 2.6;
  const color = c.P.accentTxt;
  const heroParas = c.para(c.f('chiffre', 'chiffre'), { wt: 'semi', font: 'titre', color, grad });
  const libParas = c.para(c.f('libelle', 'libellé'), { wt: 'semi', font: 'texte', g: 'libelle' });
  const oneLine = c.fit({ paras: libParas, w: S.libelleW, size: S.libelle[0], min: S.libelle[1], maxLines: 1, wt: 'semi', font: 'texte', lh: K.titre }) !== null;
  const hero = centerBox(c, { x: M, w: FULL_W, cy: oneLine ? S.cy : S.cyDeuxLignes, paras: heroParas, stages: [{ w: S.largeur, size: S.chiffre[0], min: S.chiffre[1], maxLines: 1 }], noWrap: true, wt: 'semi', align: 'center', font: 'titre', color, lh: K.titre, name: 'Chiffre', g: 'chiffre' });
  if (grad) {
    // Le dégradé suit la boîte : on la resserre sur le chiffre pour qu'il le parcoure en entier.
    const p = hero.it.paras[0];
    const ps = paraSize(hero.it, p, hero.it.fitted);
    const est = textWidth(p.runs.map((r) => r.t).join(''), ps, true, paraTrack(hero.it, p, ps), 'titre');
    const w = Math.min(FULL_W, est * 1.04 + 0.1);
    hero.it.x = (W - w) / 2;
    hero.it.w = w;
    hero.it.grad = true;
  }
  const lib = oneLine
    ? centerBox(c, { x: S.libelleX, w: S.libelleW, cy: S.libelleCy, paras: libParas, size: S.libelle[0], min: S.libelle[1], maxLines: 1, wt: 'semi', align: 'center', font: 'texte', lh: K.titre, name: 'Libellé', g: 'libelle' })
    : topBox(c, { x: S.libelleX, w: S.libelleW, capTop: S.libelleHautDeuxLignes, paras: libParas, size: S.libelle[0], min: S.libelle[1], maxTotal: 2, wt: 'semi', align: 'center', font: 'texte', lh: K.titre, balance: true, name: 'Libellé', g: 'libelle' });
  // Surtitre au corps du libellé, ligne de base à 0,25 cap du chiffre au-dessus de lui.
  c.kicker(hero.g, { x: M, w: FULL_W, align: 'center', size: S.surtitre, ecart: S.surtitreEcart });
  c.sourceLine({ y: lib.g.base + 0.2, x: S.libelleX, w: S.libelleW, g: 'libelle' });
};

// A6 : barres horizontales fines à bouts ronds ; libellé au départ (colonne alignée à droite),
// étiquette au bout ; soi en accent, l'autre en gris ; aucun axe chiffré.
R.A6 = (c) => {
  const S = STYLE.A6;
  c.accentUsed = true;
  const hasSur = !!(c.t.surtitre && String(c.t.surtitre).trim());
  const tParas = c.para(c.f('titre', 'titre du graphique'), { wt: 'semi', font: 'titre', g: 'titre' });
  const t = topBox(c, { x: M, w: FULL_W, capTop: hasSur ? S.titreHautSurtitre : S.titreHaut, paras: tParas, size: S.titre[0], min: S.titre[1], maxTotal: 2, wt: 'semi', align: 'center', font: 'titre', lh: K.titre, balance: true, name: 'Titre' });
  c.kicker(t.g, { x: M, w: FULL_W, align: 'center' });
  let bars = arr(c.t.barres).slice(0, 6);
  if (!bars.length) bars = [{ libelle: c.f('barres', 'barres'), valeur: 1 }];
  const Z = S.zone;
  const n = bars.length;
  const pitch = Z.h / n;
  const lab = clamp(Math.round(pitch * 72 * 0.36), S.etiquette[1], S.etiquette[0]);
  const labPt = scalePt(lab, 'texte');
  const vals = bars.map((b) => Math.max(0, Number(b.valeur) || 0));
  const maxV = Math.max.apply(null, vals) || 1;
  const showValues = c.t.afficher_valeurs === undefined ? ['valeurs', 'graduee', 'exigee'].includes(c.env.pub.graduation) : !!c.t.afficher_valeurs;
  const unite = c.t.unite ? ' ' + c.t.unite : '';
  const ets = bars.map((b) => {
    const et = b.etiquette !== undefined ? String(b.etiquette) : '';
    return et || (showValues ? String(b.valeur) + unite : '');
  });
  const textW = (t) => (t ? textWidth(shownText(c, t), labPt, true, trackFor('texte', labPt), 'texte') : 0);
  const etW = Math.max.apply(null, [0].concat(ets.map(textW)));
  // Libellés au départ des barres, alignés à droite dans leur colonne ; étiquettes au bout.
  const labels = bars.map((b) => str(b.libelle) || c.f('_', 'libellé'));
  const labW = Math.min(3.0, Math.max.apply(null, labels.map(textW)) + 0.05);
  const barX = Z.x + labW + 0.25;
  const maxLen = Math.max(1, Z.x + Z.w - barX - etW - 0.2);
  bars.forEach((b, i) => {
    const hl = !!b.mise_en_avant;
    const barCy = Z.y + pitch * (i + 0.5);
    const len = Math.max(S.barre * 2, vals[i] / maxV * maxLen);
    const room = Math.max(0.5, pitch);
    centerBox(c, { x: Z.x, w: labW, cy: barCy, room, text: labels[i], size: lab, min: S.etiquette[1], maxTotal: 2, wt: 'semi', color: hl ? c.P.fg : c.P.fg2, align: 'right', font: 'texte', lh: K.titre, balance: true, name: 'Libellé barre ' + (i + 1), g: 'barres', sub: i });
    c.round({ x: barX, y: barCy - S.barre / 2, w: len, h: S.barre, fill: hl ? c.P.accent : c.P.other, radius: S.barre / 2, name: 'Barre ' + (i + 1), g: 'barres', sub: i });
    if (ets[i]) centerBox(c, { x: barX + len + 0.15, w: etW + 0.3, cy: barCy, room, text: ets[i], size: lab, min: S.etiquette[1], noWrap: true, wt: 'semi', color: hl ? c.P.fg : c.P.fg2, align: 'left', font: 'texte', lh: K.titre, name: 'Valeur barre ' + (i + 1), g: 'barres', sub: i });
  });
  const axe = c.t.axe ? String(c.t.axe) : '';
  if (axe) topBox(c, { x: barX, w: Z.x + Z.w - barX, capTop: Z.y + Z.h + 0.2, text: axe, size: S.axe, min: 12, maxLines: 1, color: c.P.fg2, align: 'left', font: 'texte', lh: K.titre, name: 'Axe' });
  c.sourceLine();
};

// Titre secondaire d'une liste (A7, A8) : gris, regular, posé une cap au-dessus des lignes.
function listHeading(c, geo, o) {
  if (!c.t.titre) return geo;
  const hd = c.line({ text: c.t.titre, x: o.x, w: o.w, base: geo.capTop - geo.capH, size: scalePt(o.size, 'texte'), wt: 'regular', color: c.P.fg2, align: o.align, font: 'texte', name: 'Titre', g: 'titre' });
  return textGeom(hd);
}

R.A7 = (c) => {
  const S = STYLE.A7;
  const lines = arr(c.t.lignes);
  if (!lines.length) lines.push(c.f('lignes', 'lignes'));
  // Pas des lignes : 1,95 cap (3 à 6 lignes), 3,1 cap (2 lignes) ; l'interligne reste celui des
  // titres, le reste du pas va en espace après la ligne.
  const pitchK = lines.length === 2 ? K.liste2 : K.liste;
  const paras = listParas(c, lines, { color: c.P.fg, activeColor: c.P.fg, idleColor: c.P.fg2, font: 'titre', gapK: pitchK - K.titre });
  const stages = [{ size: S.lignes[0], min: S.lignes[1], maxLines: 1 }, { size: S.lignes[1], min: 24, maxLines: 2 }];
  const t = centerBox(c, { x: S.x, w: S.w, cy: S.cy, paras, stages, wt: 'semi', align: 'center', font: 'titre', lh: K.titre, name: 'Liste' });
  const geo = listHeading(c, t.g, { x: S.x, w: S.w, align: 'center', size: S.titre });
  c.kicker(geo, { x: S.x, w: S.w, align: 'center' });
  c.sourceLine();
};

// A8 : pictogrammes d'accent posés sur la ligne de base, texte aligné à gauche dans un bloc centré.
R.A8 = (c) => {
  const S = STYLE.A8;
  let lines = arr(c.t.lignes).slice(0, 7);
  if (!lines.length) lines = [c.f('lignes', 'lignes')];
  const coche = c.t.coche !== false;
  c.accentUsed = coche;
  const n = lines.length;
  const lineParas = lines.map((l, i) => c.para(l, { wt: 'semi', font: 'texte', g: 'lignes', sub: i }));
  const cap0 = capIn('texte', scalePt(S.lignes[0], 'texte'));
  const avail = S.largeurMax - cap0 * (S.picto + S.ecart);
  let size = null;
  lineParas.forEach((ps) => {
    const s = c.fit({ paras: ps, w: avail, size: S.lignes[0], min: S.lignes[1], maxLines: 1, wt: 'semi', font: 'texte' });
    if (s === null) c.warn('Ligne trop longue pour la liste à coches : raccourcissez-la (5 mots au plus).');
    const v = s === null ? scalePt(S.lignes[1], 'texte') : s;
    size = size === null ? v : Math.min(size, v);
  });
  const cap = capIn('texte', size);
  const pitch = K.liste * cap;
  const d = S.picto * cap;
  const gap = S.ecart * cap;
  const tw = Math.max.apply(null, lineParas.map((ps) => wrapInfo(ps.map((p) => p.runs.map((r) => r.t).join('')).join(' '), size, true, 99, trackFor('texte', size), 'texte').widest));
  const x0 = Math.max(M, (W - (d + gap + tw)) / 2);
  const cy0 = S.cy - (n - 1) * pitch / 2;
  lineParas.forEach((ps, i) => {
    const base = cy0 + i * pitch + cap / 2;
    const iy = base - d;
    // Une ligne qui n'est qu'un [À FOURNIR] garde sa case vide : une coche pleine la donnerait pour acquise.
    const aFournir = ps.some((p) => p.runs.some((r) => r.marker)) && ps.every((p) => p.runs.every((r) => r.marker || !r.t.trim()));
    if (coche && !aFournir) {
      c.ellipse({ x: x0, y: iy, w: d, h: d, fill: c.P.accent, name: 'Coche ' + (i + 1), g: 'lignes', sub: i });
      c.check({ cx: x0 + d / 2, cy: iy + d / 2, d, color: onAccent(c.P.accent, c.T.fondSombre), name: 'Coche ' + (i + 1) + ' trait', g: 'lignes', sub: i });
    } else {
      c.ellipse({ x: x0, y: iy, w: d, h: d, line: c.P.line, lineW: 1.5, name: 'Case vide ' + (i + 1), g: 'lignes', sub: i });
    }
    c.line({ paras: ps, x: x0 + d + gap, w: Math.min(tw + 0.5, W - (x0 + d + gap) - 0.2), base, size, wt: 'semi', align: 'left', font: 'texte', name: 'Ligne ' + (i + 1), g: 'lignes', sub: i });
  });
  const geo = listHeading(c, { capTop: cy0 - cap / 2, capH: cap, firstSize: size }, { x: x0, w: W - x0 - M, align: 'left', size: S.titre });
  c.kicker(geo, { x: x0, w: W - x0 - M, align: 'left' });
  c.sourceLine();
};

// A9 : l'objet sur 56 % de la largeur, à fond perdu ; l'affirmation centrée dans la colonne libre.
function sideBySide(c, defaultLabel) {
  const S = STYLE.A9;
  const right = c.t.cote === 'droite';
  const zx = right ? W - S.objet : 0;
  const col = right ? { x: W - S.colonne.x - S.colonne.w, w: S.colonne.w } : S.colonne;
  const box = { x: zx + 0.3, y: 0.3, w: S.objet - 0.6, h: H - 0.6 };
  const phBox = (kind) => {
    // Un écran schématique reste entier : coupé par le bord, sa colonne de navigation passait
    // pour une erreur de cadrage (un objet, lui, peut sortir du cadre).
    if (kind === 'interface' || kind === 'carte') return { x: right ? zx + 0.2 : 0.5, y: 0.72, w: S.objet - 0.7, h: 4.2 };
    if (kind === 'appareil') return { x: zx, y: 0.31, w: S.objet, h: 5.0 };
    if (kind === 'objet') return { x: zx + 0.4, y: 0.6, w: S.objet - 0.8, h: 4.4 };
    return box;
  };
  c.image(box, c.slide.image || { attendu: defaultLabel }, 'contain', 'Visuel', { g: 'visuel', phBox });
  const lines = arr(c.t.lignes).slice(0, 3);
  const align = lines.length ? 'left' : 'center';
  const paras = c.para(c.f('titre', 'titre'), { wt: 'semi', font: 'titre', g: 'titre', lh: K.objet });
  if (lines.length) paras[paras.length - 1].gap = S.ecartLignes;
  // Pas de 1,8 cap entre deux lignes ; une ligne qui passe sur deux garde un interligne serré.
  lines.forEach((l, i) => c.para(l, { pt: S.lignes, wt: 'regular', font: 'texte', color: c.P.fg2, g: 'lignes', sub: i, lh: K.objet, gapK: K.courte - K.objet + 0.25 }).forEach((p) => paras.push(p)));
  const t = centerBox(c, { x: col.x, w: col.w, cy: S.cy, paras, size: S.titre[0], min: S.titre[1], wt: 'semi', align, font: 'titre', lh: K.objet, balance: true, name: 'Texte' });
  c.kicker(t.g, { x: col.x, w: col.w, align });
  c.sourceLine();
}
R.A9 = (c) => sideBySide(c, c.lang === 'en' ? 'product or screen' : 'produit ou écran');
R.A11 = (c) => {
  c.warn('A11 (schéma à bloc surligné) est rendu comme A9 : visuel à gauche, légende à droite.');
  sideBySide(c, c.lang === 'en' ? 'diagram with the highlighted block' : 'schéma avec le bloc surligné');
};

// A10 : capture nue sur 85 % de la hauteur, coins arrondis, aucun texte ; un appareil sur 5,2 in.
R.A10 = (c) => {
  const S = STYLE.A10;
  const h = S.capture * H;
  const box = { x: 0.5, y: (H - h) / 2, w: W - 1.0, h };
  const phBox = (kind) => {
    if (kind === 'appareil') return { x: 0.5, y: (H - S.appareil) / 2, w: W - 1.0, h: S.appareil };
    if (kind === 'interface' || kind === 'carte') return { x: (W - h * 1.6) / 2, y: (H - h) / 2, w: h * 1.6, h };
    return box;
  };
  c.image(box, c.slide.image || { attendu: c.lang === 'en' ? 'screenshot' : 'capture d\'écran' }, 'contain', 'Écran', { g: 'visuel', radius: S.rayon, phBox });
  c.sourceLine();
};

R.A13 = (c) => {
  const S = STYLE.A13;
  c.image({ x: (W - S.w) / 2, y: (H - S.h) / 2, w: S.w, h: S.h }, c.slide.image || { attendu: c.lang === 'en' ? 'cut-out product' : 'produit détouré' }, 'contain', 'Produit', { g: 'visuel' });
  c.sourceLine();
};

function fullBleed(c, defaultLabel) {
  const S = STYLE.A18;
  c.image({ x: 0, y: 0, w: W, h: H }, c.slide.image || { attendu: defaultLabel }, 'cover', 'Photo', { kind: 'photo', full: true });
  const tw = FULL_W - (c.t.note ? 3.6 : 0);
  if (c.t.titre || c.t.surtitre) {
    c.rect({ x: 0, y: H - 1.7, w: W, h: 1.7, fill: '000000', transparency: 50, name: 'Voile' });
    const paras = c.para(c.f('titre', 'titre'), { wt: 'semi', font: 'titre', color: 'FFFFFF', g: 'titre' });
    const t = c.text({ x: M, y: H - 1.55, w: tw, h: 1.15, paras, size: S.titre[0], min: S.titre[1], wt: 'semi', align: 'left', valign: 'bottom', color: 'FFFFFF', font: 'titre', lh: K.titre, name: 'Titre' });
    c.kicker(textGeom(t), { x: M, w: tw, align: 'left', color: c.accentWords ? 'E8E8E8' : undefined });
  }
  if (c.t.note) {
    const bw = 3.4;
    c.round({ x: W - M - bw, y: H - 0.98, w: bw, h: 0.52, fill: '000000', transparency: 45, radius: 0.08, name: 'Cartouche', g: 'note' });
    c.text({ x: W - M - bw + 0.14, y: H - 0.98, w: bw - 0.28, h: 0.52, text: c.t.note, size: S.note[0], min: S.note[1], color: 'FFFFFF', align: 'left', valign: 'middle', font: 'texte', lh: K.titre, name: 'Note', g: 'note' });
  }
  if (c.sourceOnSlide) {
    c.rect({ x: 0, y: H - 0.42, w: W, h: 0.42, fill: '000000', transparency: 45, name: 'Voile source' });
    c.text({ x: M, y: H - 0.38, w: FULL_W, h: 0.34, text: c.sourceDisplay, size: 10, min: 9, color: 'FFFFFF', align: 'left', valign: 'middle', font: 'texte', lh: K.titre, name: 'Source' });
    c.sourceDone = true;
  }
}
R.A12 = (c) => fullBleed(c, c.lang === 'en' ? 'exploded or inside view' : 'vue éclatée ou intérieure');
R.A18 = (c) => fullBleed(c, c.lang === 'en' ? 'full-bleed photo' : 'photo plein cadre');
['A22', 'A23', 'A24', 'A25'].forEach((k) => {
  R[k] = (c) => {
    c.warn(k + ' est un plan de scène (vidéo) : en slide, il est rendu comme une photo plein cadre.');
    fullBleed(c, c.lang === 'en' ? 'photo of the scene' : 'photo de la scène');
  };
});

// A14 : bento jusqu'aux bords (marge et gouttière de 0,09 in), même rayon pour toutes les tuiles,
// tuile centrale de 3,9 × 2,25 in en accent, tuiles en aplat sans bordure ni ombre.
R.A14 = (c) => {
  const S = STYLE.A14;
  let tiles = arr(c.t.tuiles).slice(0, 8);
  if (arr(c.t.tuiles).length > 8) c.warn('Bento : 8 tuiles au plus autour du titre ; les suivantes sont ignorées.');
  while (tiles.length < 4) tiles.push({ titre: c.lang === 'en' ? '[TO PROVIDE: tile]' : '[À FOURNIR : tuile]' });
  const n = tiles.length;
  const m = S.marge;
  const g = S.gouttiere;
  const reserve = c.sourceOnSlide ? 0.32 : 0;
  const GW = W - 2 * m;
  const GH = H - 2 * m - reserve;
  const side = (GW - 2 * g - S.centre.w) / 2;
  const sideH = (GH - 2 * g - S.centre.h) / 2;
  const cw = [side, S.centre.w, side];
  const rh = [sideH, S.centre.h, sideH];
  const cell = (col, row, cs, rs) => {
    cs = cs || 1;
    rs = rs || 1;
    let x = m;
    for (let i = 0; i < col; i++) x += cw[i] + g;
    let y = m;
    for (let j = 0; j < row; j++) y += rh[j] + g;
    let w = -g;
    for (let i = col; i < col + cs; i++) w += cw[i] + g;
    let h = -g;
    for (let j = row; j < row + rs; j++) h += rh[j] + g;
    return { x, y, w, h };
  };
  const LAYOUTS = {
    8: [[0, 0], [1, 0], [2, 0], [0, 1], [2, 1], [0, 2], [1, 2], [2, 2]],
    7: [[0, 0, 1, 2], [1, 0], [2, 0], [2, 1], [0, 2], [1, 2], [2, 2]],
    6: [[0, 0, 1, 2], [1, 0], [2, 0, 1, 2], [0, 2], [1, 2], [2, 2]],
    5: [[0, 0, 1, 2], [1, 0], [2, 0, 1, 2], [0, 2, 2, 1], [2, 2]],
    4: [[0, 0, 1, 3], [1, 0], [2, 0, 1, 3], [1, 2]]
  };
  const rad = (b) => Math.min(S.rayon, 0.25 * Math.min(b.w, b.h));
  const center = cell(1, 1);
  const onAcc = onAccent(c.P.accent, c.T.fondSombre);
  c.round({ x: center.x, y: center.y, w: center.w, h: center.h, fill: c.P.accent, radius: rad(center), name: 'Tuile centrale', g: 'centre' });
  const cParas = c.para(c.f('titre', 'nom'), { wt: 'semi', font: 'titre', color: onAcc, g: 'centre' });
  const ct = centerBox(c, { x: center.x + 0.2, w: center.w - 0.4, cy: center.y + center.h / 2, room: center.h - 0.3, paras: cParas, size: S.nom[0], min: S.nom[1], maxTotal: 2, wt: 'semi', align: 'center', color: onAcc, font: 'titre', lh: K.titre, balance: true, name: 'Titre central', g: 'centre' });
  c.kicker(ct.g, { x: center.x + 0.2, w: center.w - 0.4, align: 'center', color: onAcc, wt: 'semi', g: 'centre' });
  LAYOUTS[n].forEach((spec, i) => {
    const b = cell(spec[0], spec[1], spec[2], spec[3]);
    const tile = typeof tiles[i] === 'string' ? { titre: tiles[i] } : (tiles[i] || {});
    c.round({ x: b.x, y: b.y, w: b.w, h: b.h, fill: c.P.surface, radius: rad(b), name: 'Tuile ' + (i + 1), g: 'tuiles', sub: i });
    const pad = S.pad;
    let area = { x: b.x + pad, y: b.y + pad, w: b.w - 2 * pad, h: b.h - 2 * pad };
    const paras = [];
    let size = S.titre;
    const label = (s) => c.para(s, { pt: S.libelle, wt: 'regular', font: 'texte', color: c.P.fg2 }).forEach((p) => paras.push(p));
    if (tile.image) {
      const ih = area.h * 0.62;
      c.image({ x: area.x, y: area.y, w: area.w, h: ih }, tile.image, 'contain', 'Visuel tuile ' + (i + 1), { g: 'tuiles', sub: i });
      area = { x: area.x, y: area.y + ih + 0.06, w: area.w, h: area.h - ih - 0.06 };
    }
    if (tile.chiffre) {
      size = S.chiffre;
      c.para(tile.chiffre, { wt: 'semi', font: 'titre', color: c.P.accentTxt, gap: 0.04 }).forEach((p) => paras.push(p));
      if (tile.titre) label(tile.titre);
      if (tile.texte) label(tile.texte);
    } else if (tile.image) {
      // Sous une image, le titre devient un libellé : il se regarde, il ne se lit pas.
      size = [S.libelle, 11];
      label(tile.titre || (c.lang === 'en' ? '[TO PROVIDE: tile]' : '[À FOURNIR : tuile]'));
      if (tile.texte) label(tile.texte);
    } else {
      c.para(tile.titre || (c.lang === 'en' ? '[TO PROVIDE: tile]' : '[À FOURNIR : tuile]'), { wt: 'semi', font: 'texte', color: c.P.fg, gap: tile.texte ? 0.06 : 0 }).forEach((p) => paras.push(p));
      if (tile.texte) label(tile.texte);
    }
    c.text({ x: area.x, y: area.y, w: area.w, h: area.h, paras, size: size[0], min: size[1], maxTotal: 4, wt: 'semi', align: 'left', valign: 'bottom', font: 'texte', lh: K.titre, balance: !tile.chiffre, name: 'Texte tuile ' + (i + 1), g: 'tuiles', sub: i });
  });
  c.sourceLine({ y: H - m - reserve + 0.05, h: 0.3 });
};

R.A15 = (c) => {
  let items = arr(c.t.elements);
  if (!items.length) items = [c.f('elements', 'éléments du mur')];
  const cols = 6;
  const rows = 5;
  const tw = 1.72;
  const th = 1.22;
  const g = 0.12;
  const x0 = -0.38;
  const y0 = -0.22;
  let k = 0;
  for (let r = 0; r < rows; r++) {
    for (let col = 0; col < cols; col++) {
      const x = x0 + col * (tw + g) + (r % 2 ? 0.45 : 0);
      const y = y0 + r * (th + g);
      const it = items[k % items.length];
      k += 1;
      const label = typeof it === 'string' ? it : (it.titre || '');
      if (it && typeof it === 'object' && it.image) {
        c.image({ x, y, w: tw, h: th }, it.image, 'cover', 'Mur ' + k, { kind: 'tuile' });
      } else {
        c.round({ x, y, w: tw, h: th, fill: c.P.surface, radius: 0.12, name: 'Mur ' + k });
      }
      if (label) c.text({ x: x + 0.12, y: y + th - 0.5, w: tw - 0.24, h: 0.4, text: label, size: STYLE.A15.libelle[0], min: STYLE.A15.libelle[1], color: c.P.fg2, align: 'left', valign: 'bottom', font: 'texte', lh: K.titre, name: 'Mur libellé ' + k });
    }
  }
};

// A16 : logos monochromes hauts de 0,56 in au plus, rangées tous les 1,1 in, grille de 8 in centrée.
R.A16 = (c) => {
  const S = STYLE.A16;
  let logos = arr(c.t.logos).slice(0, 15);
  if (!logos.length) logos = [{ nom: c.f('logos', 'logos') }];
  const n = logos.length;
  const rows = Math.ceil(n / 5);
  const cols = Math.ceil(n / rows);
  const cellW = S.largeur / cols;
  const head = !!c.t.titre;
  const cy = 2.81 + (head ? 0.25 : 0);
  const firstCy = cy - (rows - 1) * S.rangee / 2;
  logos.forEach((lg, i) => {
    const logo = typeof lg === 'string' ? { nom: lg } : lg;
    const r = Math.floor(i / cols);
    const col = i % cols;
    const inRow = Math.min(cols, n - r * cols);
    const x = (W - S.largeur) / 2 + (cols - inRow) * cellW / 2 + col * cellW;
    const box = { x: x + 0.2, y: firstCy + r * S.rangee - S.logoH / 2, w: cellW - 0.4, h: S.logoH };
    c.image(box, logo.image ? { chemin: logo.image, attendu: logo.nom } : { attendu: logo.nom || 'logo' }, 'contain', 'Logo ' + (i + 1), { kind: 'logo', g: 'logos', sub: i });
  });
  const topCap = firstCy - S.logoH / 2;
  let geo = { capTop: topCap, capH: capIn('texte', scalePt(S.titre, 'texte')) * 1.4, firstSize: scalePt(S.titre, 'texte') };
  if (head) {
    const hd = c.line({ text: c.t.titre, x: M, w: FULL_W, base: topCap - 0.5, size: scalePt(S.titre, 'texte'), wt: 'regular', color: c.P.fg2, align: 'center', font: 'texte', name: 'Titre', g: 'titre' });
    geo = textGeom(hd);
  }
  c.kicker(geo, { x: M, w: FULL_W, align: 'center' });
  c.sourceLine();
};

// A17 : citation alignée à gauche à 0,92 in, source grise 0,53 in sous la dernière ligne de base,
// bloc centré sur le milieu.
R.A17 = (c) => {
  const S = STYLE.A17;
  let q = c.f('citation', 'citation').trim();
  if (!/^[«"“]/.test(q) && !hasFournir(q)) q = c.lang === 'fr' ? '« ' + q + ' »' : '“' + q + '”';
  const qt = topBox(c, { x: S.x, w: S.w, capTop: 0, paras: c.para(q, { wt: 'semi', font: 'titre', g: 'citation' }), size: S.citation[0], min: S.citation[1], maxTotal: 4, wt: 'semi', align: 'left', font: 'titre', lh: K.citation, balance: true, name: 'Citation' });
  const at = topBox(c, { x: S.x, w: S.w, capTop: 0, text: c.f('auteur', 'auteur'), size: S.auteur[0], min: S.auteur[1], maxTotal: 2, color: c.P.fg2, align: 'left', font: 'texte', lh: K.titre, balance: true, name: 'Auteur', g: 'auteur' });
  const total = (qt.g.base - qt.g.capTop) + S.ecart + (at.g.base - at.g.capTop);
  const q2 = placeTop(qt.it, S.cy - total / 2);
  placeTop(at.it, q2.g.base + S.ecart);
  c.kicker(q2.g, { x: S.x, w: S.w, align: 'left' });
  c.sourceLine({ x: S.x, w: S.w, align: 'left' });
};

// A19 : une offre en pile (nom, petit « À partir de », prix, précision), posée à partir du haut
// des capitales du nom ; écarts mesurés. Retourne les éléments créés et la géométrie du nom.
function offerStack(c, o, x, w, align, capTop, sub, ecartNom, big) {
  const S = Object.assign({}, STYLE.A19, big ? STYLE.A19.sansVisuel : {});
  const start = c.items.length;
  const name = topBox(c, { x, w, capTop, paras: c.para(str(o.nom) || c.f('_', 'nom'), { wt: 'semi', font: 'titre', g: 'nom', sub }), size: S.nom[0], min: S.nom[1], maxTotal: 2, wt: 'semi', align, font: 'titre', lh: K.titre, balance: true, name: 'Nom', g: 'nom', sub });
  let base = name.g.base;
  let ecart = ecartNom;
  if (o.mention) {
    const m = topBox(c, { x, w, capTop: base + ecart, text: o.mention, size: S.petit, min: 12, maxLines: 1, color: c.P.fg2, align, font: 'texte', lh: K.titre, name: 'Mention', g: 'prix', sub });
    base = m.g.base;
    ecart = S.ecartPetit;
  }
  const p = topBox(c, { x, w, capTop: base + ecart, paras: c.para(str(o.prix) || c.f('_', 'prix'), { wt: 'semi', font: 'titre', g: 'prix', sub }), size: S.prix[0], min: S.prix[1], maxLines: 1, wt: 'semi', align, font: 'titre', lh: K.titre, name: 'Prix', g: 'prix', sub });
  base = p.g.base;
  if (o.note) base = topBox(c, { x, w, capTop: base + S.ecartPetit, text: o.note, size: S.petit, min: 12, maxTotal: 2, color: c.P.fg2, align, font: 'texte', lh: K.titre, name: 'Précision', g: 'prix', sub }).g.base;
  return { items: c.items.slice(start), nameItem: name.it, capTop, base, h: base - capTop };
}
// Prix seul d'un service : le prix au corps du chiffre héros (A5), le nom en haut, la période dessous.
function priceHero(c, o) {
  const S = STYLE.A19;
  const A5 = STYLE.A5;
  const hero = centerBox(c, { x: M, w: FULL_W, cy: A5.cy, paras: c.para(str(o.prix) || c.f('_', 'prix'), { wt: 'semi', font: 'titre', g: 'prix' }), stages: [{ w: A5.largeur, size: S.heros[0], min: S.heros[1], maxLines: 1 }], noWrap: true, wt: 'semi', align: 'center', font: 'titre', lh: K.titre, name: 'Prix', g: 'prix' });
  let above = hero.g;
  let ecart = A5.surtitreEcart * hero.g.capH;
  if (o.mention) {
    const m = c.line({ text: o.mention, x: M, w: FULL_W, base: hero.g.capTop - ecart, size: scalePt(A5.surtitre, 'texte'), wt: 'regular', color: c.P.fg2, align: 'center', font: 'texte', name: 'Mention', g: 'prix' });
    above = textGeom(m);
    ecart = S.ecartNomHaut;
  }
  const nm = topBox(c, { x: M, w: FULL_W, capTop: 0, paras: c.para(str(o.nom) || c.f('_', 'nom'), { wt: 'semi', font: 'titre', g: 'nom' }), size: S.nom[0], min: S.nom[1], maxTotal: 2, wt: 'semi', align: 'center', font: 'titre', lh: K.titre, balance: true, name: 'Nom', g: 'nom' });
  const g = placeTop(nm.it, above.capTop - ecart - (nm.g.base - nm.g.capTop)).g;
  if (o.note) centerBox(c, { x: A5.libelleX, w: A5.libelleW, cy: A5.libelleCy, text: o.note, size: S.periode, min: 14, maxLines: 1, color: c.P.fg2, align: 'center', font: 'texte', lh: K.titre, name: 'Précision', g: 'prix' });
  return g;
}
R.A19 = (c) => {
  const S = STYLE.A19;
  let offres = arr(c.t.offres).slice(0, 2);
  if (!offres.length) offres = [{ nom: c.t.titre, mention: c.t.mention, prix: c.t.prix, note: c.t.note }];
  const img = c.slide.image;
  const spec = c.imageSpec(img);
  const info = spec && spec.chemin ? imageInfo(path.resolve(c.env.deckDir, spec.chemin)) : null;
  if (offres.length === 1 && !img) {
    c.kicker(priceHero(c, offres[0]), { x: M, w: FULL_W, align: 'center' });
  } else if (offres.length === 1 && info && info.w / info.h >= 1.3) {
    // « Haut » : le nom en haut, l'objet dessous, sur 6 in au plus.
    const st = offerStack(c, offres[0], M, FULL_W, 'center', S.nomHaut, undefined, S.ecartNomHaut);
    const top = Math.max(S.objetHaut.y, st.base + 0.3);
    c.image({ x: (W - S.objetHaut.w) / 2, y: top, w: S.objetHaut.w, h: S.objetHaut.bas - top }, img, 'contain', 'Produit', { g: 'visuel' });
    c.kicker(textGeom(st.nameItem), { x: M, w: FULL_W, align: 'center' });
  } else if (img || offres.length === 1) {
    // « Colonne droite » : l'objet à gauche, les offres empilées et centrées à 7 in.
    // L'objet tient 91 % de la hauteur, centré vers 25 % de la largeur ; sa largeur suit ses
    // proportions (de 2,8 à 4,2 in), pour qu'un objet trapu remplisse sa moitié comme un objet élancé.
    const O = S.objetColonne;
    const ow = clamp(info ? O.h * info.w / info.h : 3.7, O.w, 4.2);
    c.image({ x: O.cx - ow / 2, y: O.y, w: ow, h: O.h }, img, 'contain', 'Produit', { g: 'visuel' });
    const x = S.colonne.cx - S.colonne.w / 2;
    const stacks = [];
    let cap = 0;
    offres.forEach((o, i) => {
      const st = offerStack(c, o, x, S.colonne.w, 'center', cap, offres.length > 1 ? i : undefined, S.ecartNomColonne);
      stacks.push(st);
      cap = st.base + S.ecartOffres;
    });
    const total = stacks[stacks.length - 1].base;
    stacks.forEach((st) => shiftItems(st.items, S.cy - total / 2));
    c.kicker(textGeom(stacks[0].nameItem), { x, w: S.colonne.w, align: 'center' });
  } else {
    // Deux offres sans visuel : côte à côte, alignées sur leur nom, le bloc centré sur le milieu.
    const stacks = offres.map((o, i) => offerStack(c, o, S.deuxColonnes[i] - S.colonne.w / 2, S.colonne.w, 'center', 0, i, S.ecartNomColonne, true));
    const total = Math.max.apply(null, stacks.map((st) => st.base));
    stacks.forEach((st) => shiftItems(st.items, S.cy - total / 2));
    c.kicker(textGeom(stacks[0].nameItem), { x: M, w: FULL_W, align: 'center', ecart: 1.4 });
  }
  c.sourceLine();
};

// A20 : petit libellé gris en haut, le mot ou la date dessous, le visuel calé en bas ; sans visuel,
// libellé de 24 pt et mot au corps de la carte-titre, le bloc centré sur le milieu.
R.A20 = (c) => {
  const S = STYLE.A20;
  const img = c.slide.image;
  const label = c.t.libelle ? String(c.t.libelle) : '';
  const motParas = c.para(c.f('titre', 'date'), { wt: 'semi', font: 'titre', g: 'date' });
  let geo;
  if (img) {
    let base = null;
    if (label) {
      const l = topBox(c, { x: M, w: FULL_W, capTop: S.libelleHaut, text: label, size: S.libelle, min: 12, maxLines: 1, color: c.P.fg2, align: 'center', font: 'texte', lh: K.titre, name: 'Libellé', g: 'libelle' });
      base = l.g.base;
      geo = l.g;
    }
    const m = topBox(c, { x: M, w: FULL_W, capTop: base === null ? S.libelleHaut : base + S.ecart, paras: motParas, size: S.mot[0], min: S.mot[1], maxTotal: 2, wt: 'semi', align: 'center', font: 'titre', lh: K.titre, balance: true, name: 'Date', g: 'date' });
    if (!geo) geo = m.g;
    const V = S.visuel;
    const top = Math.max(V.y, m.g.base + 0.3);
    c.image({ x: (W - V.w) / 2, y: top, w: V.w, h: V.bas - top }, img, 'contain', 'Produit', { g: 'visuel', alignY: 'bottom', phBox: () => ({ x: (W - 6.4) / 2, y: top + 0.1, w: 6.4, h: V.bas - top - 0.1 }) });
  } else {
    const A = STYLE.A1;
    const m = topBox(c, { x: A.x, w: A.w, capTop: 0, paras: motParas, stages: bigStages(A.corps[0], motParas.length > 1), wt: 'semi', align: 'center', font: 'titre', lh: K.titre, balance: true, name: 'Date', g: 'date' });
    const l = label ? topBox(c, { x: M, w: FULL_W, capTop: 0, text: label, size: S.seul.libelle, min: 16, maxLines: 1, color: c.P.fg2, align: 'center', font: 'texte', lh: K.titre, name: 'Libellé', g: 'libelle' }) : null;
    const gap = l ? S.seul.ecart * m.g.capH : 0;
    const lh = l ? l.g.base - l.g.capTop : 0;
    const top = A.cy - (lh + gap + (m.g.base - m.g.capTop)) / 2;
    if (l) {
      geo = placeTop(l.it, top).g;
      placeTop(m.it, geo.base + gap);
    } else {
      geo = placeTop(m.it, top).g;
    }
  }
  c.kicker(geo, { x: M, w: FULL_W, align: 'center', ecart: 0.9 });
  c.sourceLine();
};

// A21 : nom et fonction au même corps, en bas à gauche, en blanc sur l'image, sans fond ni ombre.
R.A21 = (c) => {
  const S = STYLE.A21;
  const who = c.env.people[c.t.intervenant || c.slide.intervenant] || {};
  const nom = c.t.nom || who.nom || c.f('nom', 'nom');
  const fonction = c.t.fonction || who.role || c.f('fonction', 'fonction');
  const img = c.slide.image;
  if (img) c.image({ x: 0, y: 0, w: W, h: H }, img, 'cover', 'Photo', { kind: 'photo', full: true });
  const size = scalePt(S.corps, 'texte');
  c.line({ text: nom, x: S.x, w: 6, base: S.bases[0], size, wt: 'semi', color: img ? 'FFFFFF' : c.P.fg, align: 'left', font: 'texte', name: 'Nom', g: 'nom' });
  c.line({ text: fonction, x: S.x, w: 6, base: S.bases[1], size, wt: 'regular', color: img ? 'FFFFFF' : c.P.fg2, align: 'left', font: 'texte', name: 'Fonction', g: 'fonction' });
};

// AF : slide entière à fournir. Avec archetype_prevu, la slide est rendue dans cet archétype
// (les champs manquants deviennent des [À FOURNIR] courts) ; sinon, un libellé centré.
R.AF = (c) => {
  let label = str(c.t.titre || c.t.libelle).trim();
  if (!label) label = c.lang === 'en' ? '[TO PROVIDE: ' + (c.hint() || c.L.content) + ']' : '[À FOURNIR : ' + (c.hint() || c.L.content) + ']';
  else if (!hasFournir(label)) label = c.lang === 'en' ? '[TO PROVIDE: ' + label + ']' : '[À FOURNIR : ' + label + ']';
  const S = STYLE.A4;
  const t = centerBox(c, { x: S.x, w: S.w, cy: S.cy, paras: c.para(label, { wt: 'semi', font: 'titre', g: 'titre' }), size: STYLE.AF.titre[0], min: STYLE.AF.titre[1], wt: 'semi', align: 'center', font: 'titre', lh: K.titre, balance: true, name: 'Titre' });
  c.kicker(t.g, { x: S.x, w: S.w, align: 'center' });
};

R.SOURCES = (c) => {
  const S = STYLE.SOURCES;
  topBox(c, { x: M, w: FULL_W, capTop: 0.6, text: c.L.sourcesTitle, size: S.titre[0], min: S.titre[1], maxLines: 1, wt: 'semi', color: c.P.fg2, align: 'left', font: 'texte', lh: K.titre, name: 'Titre' });
  const paras = [];
  c.env.collectedSources.forEach((s) => {
    const line = s.n + '. ' + s.text + ' (' + c.L.slideWord + (s.slides.length > 1 ? 's ' : ' ') + s.slides.join(', ') + ')';
    c.para(line, { color: c.P.fg2, gap: 0.1 }).forEach((p) => paras.push(p));
  });
  c.text({ x: M, y: 1.25, w: FULL_W, h: H - 1.25 - 0.5, paras, size: S.liste[0], min: S.liste[1], align: 'left', valign: 'top', font: 'texte', lh: K.titre, name: 'Sources' });
};

// [À FOURNIR] en grand corps (au-dessus de STYLE.marqueur.seuil) : le soulignement de la police
// deviendrait une rangée de gros points. On le retire du texte (r.ul = false) et on pose, juste
// avant la zone de texte (donc dessous), un filet pointillé d'épaisseur fixe sous la portion du
// marqueur. Seulement pour un paragraphe d'une ligne ; au-delà, le soulignement reste celui du texte.
function markerRules(c) {
  const MK = STYLE.marqueur;
  const inserts = [];
  c.items.forEach((it, idx) => {
    if (it.k !== 'text' || !it.paras.some((p) => p.runs.some((r) => r.marker))) return;
    const geo = textGeom(it);
    let li = 0;
    it.paras.forEach((p) => {
      const ps = paraSize(it, p, it.fitted);
      const n = it.noWrap ? 1 : paraInfo(it, p, ps).lines;
      const line = geo.lines[li];
      li += n;
      if (n !== 1 || !line || ps < MK.seuil || !p.runs.some((r) => r.marker)) return;
      const hv = heavy(paraWeight(it, p));
      const tr = paraTrack(it, p, ps);
      const role = roleOf(it, p);
      // Position et largeur estimées pour chaque sortie avec la police qu'elle affiche : le .pptx avec
      // la police écrite (cle), l'aperçu avec celle qu'il montre ici (ici, le repli si elle manque).
      // k corrige la table des chasses, mesuré dans l'aperçu (rapport réel / estimé).
      const k = hv ? MK.chasse[0] : MK.chasse[1];
      const F = FONT[role] || FONT.texte;
      const facteur = (key) => (STYLE.polices[key] || STYLE.polices.defaut).largeur;
      const align = p.align || it.align;
      const spans = (wf) => {
        const wOf = (s) => textWidth(s, ps, hv, tr, role, wf) * k;
        const lw = wOf(p.runs.map((r) => r.t).join(''));
        const x0 = align === 'center' ? it.x + (it.w - lw) / 2 : align === 'right' ? it.x + it.w - lw : it.x;
        const out = [];
        let acc = 0;
        p.runs.forEach((r) => {
          const rw = wOf(r.t);
          if (r.marker) {
            const lead = wOf(r.t.match(/^\s*/)[0]);
            const trail = wOf(r.t.match(/\s*$/)[0]);
            out.push({ x: x0 + acc + lead, w: Math.max(0.05, rw - lead - trail) });
          }
          acc += rw;
        });
        return out;
      };
      const ecrite = spans(facteur(F.cle));
      const vue = spans(facteur(F.ici || F.cle));
      const thick = MK.epaisseur / 72;
      const y = line.c + line.cap / 2 + clamp(ps * MK.ecart, MK.ecartPt[0], MK.ecartPt[1]) / 72;
      let mi = 0;
      p.runs.forEach((r) => {
        if (r.marker) {
          r.ul = false;
          const a = ecrite[mi];
          const b = vue[mi];
          mi += 1;
          inserts.push({ before: idx, it: { k: 'rule', x: a.x, y: y - thick / 2, w: a.w, xh: b.x, wh: b.w, h: thick, pt: MK.epaisseur, color: c.P.accentTxt, name: 'Soulignement', g: p.g !== undefined ? p.g : it.g, sub: p.sub !== undefined ? p.sub : it.sub, ph: it.ph } });
        }
      });
    });
  });
  inserts.sort((a, b) => b.before - a.before).forEach((ins) => {
    const n = (c.names[ins.it.name] || 0) + 1;
    c.names[ins.it.name] = n;
    if (n > 1) ins.it.name += ' ' + n;
    if (ins.it.ph === undefined) delete ins.it.ph;
    c.items.splice(ins.before, 0, ins.it);
  });
}

// ---------------------------------------------------------------------------
// Composition d'une slide (fond + archétype + logo)
// ---------------------------------------------------------------------------
function composeSlide(deck, slide, index, env) {
  const c = new SlideCtx(deck, slide, index, env);
  let fn = R[c.arch];
  if (c.arch === 'AF') {
    const prevu = String(c.t.archetype_prevu || slide.archetype_prevu || '').toUpperCase();
    if (prevu && prevu !== 'AF' && R[prevu]) {
      c.renderArch = prevu;
      fn = R[prevu];
    }
  }
  if (c.fond === 'image' && !FULL_BLEED.has(c.renderArch)) {
    const bgSpec = slide.image_fond || (!CONTENT_IMAGE.has(c.renderArch) ? slide.image : null);
    if (bgSpec) {
      c.image({ x: 0, y: 0, w: W, h: H }, bgSpec, 'cover', 'Image de fond', { kind: 'photo', full: true });
      const voile = slide.voile === undefined ? 45 : Math.max(0, Math.min(100, Number(slide.voile) || 0));
      if (voile > 0) c.rect({ x: 0, y: 0, w: W, h: H, fill: '000000', transparency: 100 - voile, name: 'Voile' });
    } else {
      c.warn('Fond « image » sans image_fond : fond sombre utilisé.');
    }
  }
  if (fn) {
    fn(c);
  } else {
    c.warn('Archétype inconnu « ' + slide.archetype + ' » : libellé [À FOURNIR] à la place.');
    c.renderArch = 'AF';
    c.t = { titre: c.L.content };
    R.AF(c);
  }
  markerRules(c);
  if (slide.logo && env.T.logo) {
    const abs = path.resolve(env.deckDir, env.T.logo);
    const info = imageInfo(abs);
    if (info) {
      const s = Math.min(1.2 / info.w, 0.38 / info.h);
      c.add({ k: 'image', x: W - 0.45 - info.w * s, y: H - 0.42 - info.h * s, w: info.w * s, h: info.h * s, abs, info, mode: 'fill', name: 'Logo', alt: 'logo' });
    } else c.warn('Logo « ' + env.T.logo + ' » introuvable : non affiché.');
  }
  return c;
}

// ---------------------------------------------------------------------------
// Apparitions : slide.build (phrases pour l'orateur) → animations d'entrée en fondu, au clic.
// Chaque archétype a ses groupes, dans l'ordre de lecture. Une phrase qui nomme un groupe
// (« le chiffre », « les trois lignes une à une », « la tuile centrale ») le fait apparaître ;
// une phrase qui dit « déjà à l'écran » le laisse affiché ; « ensemble » regroupe ; « puis »
// sépare ; « suit » enchaîne sans clic. Un groupe jamais nommé apparaît avec le groupe qui
// le précède, ou reste affiché s'il vient en premier. Sans build, rien n'est animé.
// ---------------------------------------------------------------------------
const KW = {
  surtitre: /\b(?:surtitres?|sur-titres?|kickers?|eyebrows?|overlines?)\b/,
  titre: /\b(?:titres?|titles?|noms?|names?|phrases?|sentences?|mots?|words?|headings?|headlines?|accroches?|slogans?|statements?|messages?|claims?)\b/,
  ligne: /\b(?:lignes?|lines?)\b/,
  lignes: /\b(?:lignes?|lines?|points?|items?|elements?|coches?|cases?|checks?|ticks?|checkmarks?|garanties?|atouts?|benefices?|bullets?|rows?|etapes?|steps?)\b/,
  chiffre: /\b(?:chiffres?|nombres?|numbers?|figures?|valeurs?|values?|stats?|statistiques?|pourcentages?|percentages?|gains?)\b/,
  libelle: /\b(?:libelles?|labels?|legendes?|captions?|sous-titres?|subtitles?|explications?|textes?|texts?|phrases?|sentences?)\b/,
  visuel: /\b(?:captures?|ecrans?|images?|visuels?|photos?|produits?|boitiers?|objets?|schemas?|illustrations?|screens?|screenshots?|visuals?|pictures?|products?|devices?|appareils?|maquettes?|mockups?|icones?|icons?|logos?|diagrammes?|diagrams?|interfaces?|cartes?|maps?)\b/,
  barres: /\b(?:barres?|bars?|colonnes?|columns?)\b/,
  tuiles: /\b(?:tuiles?|tiles?|blocs?|blocks?|cartes?|cards?|cases?)\b/,
  centre: /\b(?:tuile centrale|central tile|centre tile|center tile|centre|center|centrale|central|middle)\b/,
  prix: /\b(?:prix|prices?|tarifs?|montants?|amounts?|pricing)\b/,
  nom: /\b(?:noms?|names?|offres?|offers?|titres?|titles?|formules?|plans?)\b/,
  date: /\b(?:dates?|jours?|days?|mois|months?|titres?|titles?|mots?|words?)\b/,
  citation: /\b(?:citations?|quotes?|quotations?|temoignages?|testimonials?|phrases?|sentences?)\b/,
  auteur: /\b(?:auteurs?|authors?|signatures?|attributions?)\b/,
  logos: /\b(?:logos?|marques?|brands?|partenaires?|partners?|clients?)\b/,
  note: /\b(?:notes?|cartouches?|mentions?|legendes?|captions?)\b/,
  fonction: /\b(?:fonctions?|roles?|postes?|jobs?)\b/
};
const BRX = {
  none: /\b(?:aucune apparition|pas d'apparition|sans apparition|tout (?:est )?a l'ecran d'un bloc|no build|no animation|all at once on screen)\b/,
  seq: /\b(?:puis|ensuite|then|followed by|apres quoi)\b/,
  stat: /\bdeja\b|\balready\b|des l'arrivee|des le debut|avant (?:la prise de parole|de parler|qu'(?:il|elle) parle)|before (?:speaking|the speaker|anyone speaks)|from the start/,
  keep: /\b(?:reste|restent|stays?|remains?)\b/,
  together: /\b(?:ensemble|en meme temps|simultanement|d'un coup|together|at once|at the same time|simultaneously)\b/,
  after: /\b(?:suit|suivent|juste apres|une seconde apres|aussitot apres|follows?|right after|a second later|shortly after)\b/,
  appear: /(?:apparai|apparu|s'affiche|affichent|arrive|s'allume|allument|se pose|se posent|surgit|appears?|shows? up|comes? in|lights? up|fades? in|pops? in|builds? in)/,
  plural: /\b(?:toutes|tous|chaque|une a une|un a un|l'une apres l'autre|l'un apres l'autre|all|each|one by one|one after another|in turn)\b/
};
const ORD = [
  [/\b(?:premiere|premier|1re|1ere|1er|first)\b/g, 0],
  [/(?<!\b(?:une|un|a|one) )\b(?:seconde|second|deuxieme|2e|2nd)\b/g, 1],
  [/\b(?:troisieme|3e|third|3rd)\b/g, 2],
  [/\b(?:quatrieme|4e|fourth|4th)\b/g, 3],
  [/\b(?:cinquieme|5e|fifth|5th)\b/g, 4],
  [/\b(?:sixieme|6e|sixth|6th)\b/g, 5],
  [/\b(?:derniere|dernier|last)\b/g, 'last']
];
const G = (key, kw, o) => Object.assign({ key, kw }, o || {});
const SUR = G('surtitre', KW.surtitre, { prio: 1 });
const GROUPS = {
  A1: [SUR, G('titre', KW.titre)],
  A2: [G('visuel', KW.visuel), SUR, G('titre', KW.titre)],
  A3: [SUR, G('lignes', KW.lignes, { multi: true })],
  A4: [SUR, G('titre', KW.titre, { subKw: KW.ligne })],
  A5: [SUR, G('chiffre', KW.chiffre), G('libelle', KW.libelle)],
  A6: [SUR, G('titre', KW.titre), G('barres', KW.barres, { multi: true })],
  A7: [SUR, G('titre', KW.titre), G('lignes', KW.lignes, { multi: true })],
  A8: [SUR, G('titre', KW.titre), G('lignes', KW.lignes, { multi: true })],
  A9: [G('visuel', KW.visuel), SUR, G('titre', KW.titre), G('lignes', KW.lignes, { multi: true })],
  A12: [SUR, G('titre', KW.titre), G('note', KW.note)],
  A14: [G('centre', KW.centre, { prio: 1 }), G('tuiles', KW.tuiles, { multi: true })],
  A16: [SUR, G('titre', KW.titre), G('logos', KW.logos, { multi: true })],
  A17: [SUR, G('citation', KW.citation), G('auteur', KW.auteur)],
  A19: [SUR, G('visuel', KW.visuel), G('nom', KW.nom, { multi: 'auto' }), G('prix', KW.prix, { multi: 'auto' })],
  A20: [SUR, G('libelle', KW.libelle), G('date', KW.date), G('visuel', KW.visuel)],
  A21: [G('nom', KW.titre), G('fonction', KW.fonction)],
  AF: [SUR, G('titre', KW.titre)]
};
GROUPS.A11 = GROUPS.A9;
['A18', 'A22', 'A23', 'A24', 'A25'].forEach((k) => { GROUPS[k] = GROUPS.A12; });

// Cibles d'un groupe : éléments entiers (it.g) ou plages de paragraphes d'une zone de texte.
function groupTargets(c, key) {
  const bySub = {};
  const all = [];
  const push = (sub, t) => {
    all.push(t);
    if (sub !== undefined) (bySub[sub] = bySub[sub] || []).push(t);
  };
  c.items.forEach((it, idx) => {
    if (it.g === key) {
      push(it.sub, { idx, pr: null });
    } else if (it.k === 'text' && it.g === undefined) {
      let start = -1;
      let cur;
      const flush = (end) => {
        if (start >= 0) push(cur, { idx, pr: [start, end] });
        start = -1;
      };
      it.paras.forEach((p, pi) => {
        if (p.g === key) {
          if (start >= 0 && p.sub !== cur) flush(pi - 1);
          if (start < 0) {
            start = pi;
            cur = p.sub;
          }
        } else {
          flush(pi - 1);
        }
      });
      flush(it.paras.length - 1);
    }
  });
  const subs = Object.keys(bySub).map(Number).sort((a, b) => a - b);
  return { all, bySub, subs };
}
function ordinalsIn(seg) {
  const found = [];
  ORD.forEach(([rx, v]) => {
    rx.lastIndex = 0;
    let m;
    while ((m = rx.exec(seg))) found.push({ pos: m.index, v });
  });
  return found.sort((a, b) => a.pos - b.pos).map((x) => x.v);
}
function isPluralWord(word, seg, pos) {
  const w = String(word || '');
  if (/s$/.test(w) && !['prix', 'mois', 'fois', 'cas'].includes(w)) return true;
  // « les prix », « les deux offres », « the three lines » : un déterminant pluriel juste avant.
  const before = seg.slice(Math.max(0, (pos || 0) - 24), pos || 0);
  return /\b(?:les|des|ces|nos|vos|the|these|those|all|both)\b(?:\s+[a-z]+)?\s+$/.test(before) || BRX.plural.test(seg);
}
function planAnimations(c) {
  const raw = arr(c.slide.build).map((b) => {
    if (typeof b === 'string') return b;
    if (b && typeof b === 'object') return [b.element, b.quoi, b.texte, b.text].filter(Boolean).join(' ');
    return '';
  }).filter((s) => s.trim());
  if (!raw.length) return [];
  const defs = [];
  (GROUPS[c.renderArch] || []).forEach((d, i) => {
    const t = groupTargets(c, d.key);
    if (!t.all.length) return;
    const multi = d.multi === true || (d.multi === 'auto' && t.subs.length > 1);
    defs.push(Object.assign({}, d, { order: i, t, multi, done: new Set(), whole: false, isStatic: false, lastStep: -1 }));
  });
  if (!defs.length) return [];
  const byPrio = defs.slice().sort((a, b) => (a.prio || 5) - (b.prio || 5) || a.order - b.order);
  const isDone = (d) => d.whole || (d.t.subs.length > 0 && d.t.subs.every((k) => d.done.has(k)));
  const restOf = (d) => {
    if (d.whole) return [];
    if (!d.t.subs.length) return d.t.all;
    const out = [];
    d.t.subs.filter((k) => !d.done.has(k)).forEach((k) => d.t.bySub[k].forEach((t) => out.push(t)));
    // Éléments sans rang (rares) : avec le reste.
    d.t.all.filter((t) => !d.t.subs.some((k) => d.t.bySub[k].includes(t))).forEach((t) => out.push(t));
    return out;
  };
  const steps = [];
  for (const entry of raw) {
    const n = normText(entry).replace(/«[^»]*»|“[^”]*”|"[^"]*"/g, ' ');
    if (BRX.none.test(n)) return [];
    const entryTogether = BRX.together.test(n);
    let lastDef = null;
    n.split(BRX.seq).forEach((seg) => {
      const isStatic = BRX.stat.test(seg);
      const keep = BRX.keep.test(seg) && !BRX.appear.test(seg);
      const after = BRX.after.test(seg);
      const together = entryTogether || BRX.together.test(seg);
      let s = seg;
      const mentions = [];
      byPrio.forEach((d) => {
        const tryRx = (rx, sub) => {
          const m = rx.exec(s);
          if (!m) return false;
          mentions.push({ d, pos: m.index, word: m[0], sub });
          s = s.slice(0, m.index) + ' '.repeat(m[0].length) + s.slice(m.index + m[0].length);
          return true;
        };
        if (d.subKw && tryRx(d.subKw, true)) return;
        tryRx(d.kw, false);
      });
      mentions.sort((a, b) => a.pos - b.pos);
      const ords = ordinalsIn(seg);
      if (!mentions.length && ords.length && lastDef) mentions.push({ d: lastDef, pos: 0, word: '', sub: true });
      if (!mentions.length && (isStatic || BRX.appear.test(seg))) {
        const d = defs.find((x) => !isDone(x));
        if (d) mentions.push({ d, pos: 0, word: '', implicit: true });
      }
      if (!mentions.length) return;
      if (keep && !isStatic) {
        lastDef = mentions[mentions.length - 1].d;
        return;
      }
      const units = [];
      mentions.forEach((mn) => {
        const d = mn.d;
        lastDef = d;
        if (isDone(d)) return;
        const subLevel = d.t.subs.length > 0 && (d.multi || mn.sub || (ords.length > 0 && d.t.subs.length > 1));
        if (subLevel) {
          const free = d.t.subs.filter((k) => !d.done.has(k));
          let chosen;
          if (ords.length) chosen = uniq(ords.map((o) => (o === 'last' ? d.t.subs[d.t.subs.length - 1] : d.t.subs[o])).filter((k) => k !== undefined && !d.done.has(k)).map(String)).map(Number);
          else if (mn.implicit || isPluralWord(mn.word, seg, mn.pos)) chosen = free;
          else chosen = free.slice(0, 1);
          if (!chosen.length) return;
          chosen.forEach((k) => d.done.add(k));
          if (together || isStatic) units.push({ d, targets: [].concat.apply([], chosen.map((k) => d.t.bySub[k])), label: [d.key, chosen], multi: false });
          else chosen.forEach((k) => units.push({ d, targets: d.t.bySub[k], label: [d.key, [k]], multi: true }));
        } else {
          const targets = restOf(d);
          d.whole = true;
          if (targets.length) units.push({ d, targets, label: [d.key, null], multi: false });
        }
      });
      if (!units.length) return;
      if (isStatic) {
        units.forEach((u) => { u.d.isStatic = true; });
        return;
      }
      units.forEach((u, j) => {
        let trigger;
        if (j === 0) trigger = after && steps.length ? 'after' : 'click';
        else trigger = u.multi && !together ? 'click' : 'with';
        steps.push({ trigger, targets: u.targets.slice(), labels: [u.label] });
        u.d.lastStep = steps.length - 1;
      });
    });
  }
  if (!steps.length) return [];
  // Groupes jamais nommés : avec le groupe animé qui les précède dans l'ordre de lecture.
  defs.forEach((d) => {
    if (d.isStatic) return;
    const rest = restOf(d);
    if (!rest.length) return;
    const prev = defs.filter((x) => x.order <= d.order && x.lastStep >= 0).pop();
    if (!prev) return;
    steps[prev.lastStep].targets.push.apply(steps[prev.lastStep].targets, rest);
    steps[prev.lastStep].labels.push([d.key, null]);
    d.whole = true;
  });
  steps[0].trigger = 'click';
  return steps;
}
function describeSteps(steps, L) {
  const parts = [];
  let k = 0;
  steps.forEach((s) => {
    const names = s.labels.map(([key, subs]) => {
      const nm = L.parts[key] || key;
      return subs ? subs.map((x) => nm + ' ' + (x + 1)).join(' + ') : nm;
    }).join(' + ');
    if (s.trigger === 'click') {
      k += 1;
      parts.push(k + ') ' + names);
    } else {
      parts[parts.length - 1] += (s.trigger === 'after' ? ' → ' : ' + ') + names;
    }
  });
  return parts.join(' · ');
}

// ---------------------------------------------------------------------------
// Notes du présentateur : le texte à dire, puis apparitions, minutage, source, ce qui
// reste à valider et à fournir. Aucun terme de travail (étape, archétype).
// ---------------------------------------------------------------------------
function noteSegments(slide, env) {
  const def = slide.intervenant || (env.intervenants[0] && env.intervenants[0].id) || '';
  const n = slide.notes;
  if (n === undefined || n === null || n === '') return [];
  if (typeof n === 'string') return [{ intervenant: def, texte: n }];
  return arr(n).map((s) => (typeof s === 'string' ? { intervenant: def, texte: s } : { intervenant: s.intervenant || def, texte: str(s.texte || s.text), type: s.type }))
    .filter((s) => s.texte.trim() !== '');
}
// En-tête d'une prise de parole : NOM · fonction. Une fonction inconnue (vide ou écrite en
// marqueur) est signalée une seule fois, à la première prise de parole de la personne.
function speakerHead(id, env, L) {
  const who = env.people[id];
  if (!who) return String(id || '').toUpperCase();
  const name = String(stripValider(who.nom || id)).toUpperCase();
  const role = str(who.role).trim();
  const first = !env.introduced.has(id);
  env.introduced.add(id);
  if (role && !hasFournir(role)) return name + ' · ' + stripValider(role);
  if (first) return name + ' · ' + (hasFournir(role) ? role : (env.lang === 'en' ? '[TO PROVIDE: ' + L.role + ']' : '[À FOURNIR : ' + L.role + ']'));
  return name;
}
function provideItems(slide, c, env) {
  const L = c.L;
  const out = [];
  if (c.fournirs.length) out.push(c.fournirs.join(', ') + '.');
  arr(slide.a_fournir).forEach((x) => { if (str(x).trim()) out.push(str(x).trim()); });
  if (c.arch === 'AF' && c.t.attendu) out.push(str(c.t.attendu).trim());
  c.missing.forEach((m) => {
    if (m.kind === 'logo' && !m.chemin) return;
    out.push((L.kind[m.kind] || L.kind.interface) + (m.attendu ? ' ' + L.q(m.attendu) : '') + (m.chemin ? ' (' + m.chemin + ')' : '') + '.');
  });
  return uniq(out);
}
function buildNotes(slide, c, env, start) {
  const L = c.L;
  const out = [];
  const segs = noteSegments(slide, env);
  if (!segs.length) {
    out.push(L.noNotes);
    c.warn('Notes vides : marqueur ajouté dans les notes.');
  }
  segs.forEach((s) => {
    const changed = env.lastSpeaker && s.intervenant && s.intervenant !== env.lastSpeaker;
    const head = (changed ? '→ ' : '') + speakerHead(s.intervenant, env, L) + (s.type === 'relais' ? ' (' + L.relay + ')' : '');
    if (s.intervenant) out.push(head);
    out.push(s.texte.trim());
    out.push('');
    if (s.intervenant) env.lastSpeaker = s.intervenant;
  });
  const build = arr(slide.build).map((b) => (typeof b === 'string' ? b : str(b && (b.texte || b.text)))).filter((b) => b.trim());
  if (build.length) {
    out.push(L.build);
    build.forEach((b, i) => out.push((i + 1) + '. ' + b));
    out.push('');
  }
  const d = Number(slide.duree_s) || 0;
  out.push(L.timing + ' ' + fmtTime(start) + ' → ' + fmtTime(start + d) + ' (' + d + ' s)');
  if (c.sourceText) {
    const where = env.sourceMode === 'renvoi' ? L.sourceRenvoi + c.sourceRef + ', ' + L.sourceFinal
      : c.sourceOnSlide ? L.sourceShown : env.sourceMode === 'carton_final' ? L.sourceFinal : L.sourceHidden;
    out.push(L.source + ' (' + where + ')' + L.sep + c.sourceText);
  }
  if (c.proposals.length) out.push(L.toValidate + L.sep + c.proposals.map((p) => L.q(p)).join(c.lang === 'en' ? '; ' : ' ; '));
  const prov = provideItems(slide, c, env);
  if (prov.length) out.push(L.toProvide + L.sep + prov.join(' '));
  return out.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

// ---------------------------------------------------------------------------
// Rendu .pptx. Retourne, pour chaque élément, les identifiants des formes créées
// (pptxgenjs numérote les objets d'une slide à partir de 2, dans l'ordre d'ajout).
// Interligne exact (a:spcPts), suivi en points (spc), demi-gras par la face SemiBold.
// ---------------------------------------------------------------------------
function shadowPt() {
  const O = STYLE.schema.ombre;
  return { type: 'outer', blur: Math.round(O.flou * 72 * 10) / 10, offset: Math.round(O.decalage * 72 * 10) / 10, angle: 90, color: '000000', opacity: O.opacite };
}
function pptxRender(pres, slideObj, c, env) {
  const T = env.T;
  const langTag = env.lang === 'en' ? 'en-US' : 'fr-FR';
  const map = [];
  let n = 0;
  const reg = (idx, type, names) => {
    map[idx] = { ids: names.map((nm, k) => n + 2 + k), names, type };
    n += names.length;
  };
  c.items.forEach((it, idx) => {
    if (it.k === 'rect' || it.k === 'round' || it.k === 'ellipse') {
      const shape = it.k === 'ellipse' ? pres.ShapeType.ellipse : it.k === 'round' ? pres.ShapeType.roundRect : pres.ShapeType.rect;
      const o = { x: round2(it.x), y: round2(it.y), w: round2(Math.max(0.01, it.w)), h: round2(Math.max(0.01, it.h)), objectName: it.name };
      if (it.fill) o.fill = it.transparency ? { color: it.fill, transparency: it.transparency } : { color: it.fill };
      if (it.line) o.line = { color: it.line, width: it.lineW || 1, dashType: it.dash || 'solid' };
      if (it.k === 'round') o.rectRadius = round2(Math.max(0, Math.min(it.radius || 0.1, it.w / 2, it.h / 2)));
      if (it.shadow) o.shadow = shadowPt();
      slideObj.addShape(shape, o);
      reg(idx, 'sp', [it.name]);
    } else if (it.k === 'check') {
      const s = it.d;
      const a = { x: it.cx - 0.24 * s, y: it.cy + 0.0 * s };
      const b = { x: it.cx - 0.06 * s, y: it.cy + 0.18 * s };
      const e = { x: it.cx + 0.26 * s, y: it.cy - 0.2 * s };
      const w = Math.max(1.5, s * 7);
      slideObj.addShape(pres.ShapeType.line, { x: round2(a.x), y: round2(a.y), w: round2(b.x - a.x), h: round2(b.y - a.y), line: { color: it.color, width: w }, objectName: it.name + ' a' });
      slideObj.addShape(pres.ShapeType.line, { x: round2(b.x), y: round2(e.y), w: round2(e.x - b.x), h: round2(b.y - e.y), flipV: true, line: { color: it.color, width: w }, objectName: it.name + ' b' });
      reg(idx, 'sp', [it.name + ' a', it.name + ' b']);
    } else if (it.k === 'rule') {
      // Filet pointillé d'un [À FOURNIR] en grand corps : épaisseur fixe, opacité du marqueur.
      slideObj.addShape(pres.ShapeType.line, { x: round2(it.x), y: round2(it.y + it.h / 2), w: round2(it.w), h: 0, line: { color: it.color, width: it.pt, dashType: 'sysDot', transparency: Math.round((1 - STYLE.marqueur.opacite) * 100) }, objectName: it.name });
      reg(idx, 'sp', [it.name]);
    } else if (it.k === 'image') {
      if (it.mode === 'cover') {
        const iw = it.info.w;
        const ih = it.info.h;
        slideObj.addImage({ path: it.abs, x: round2(it.x), y: round2(it.y), w: round2(it.w), h: round2(it.w * ih / iw), sizing: { type: 'cover', w: round2(it.w), h: round2(it.h) }, altText: it.alt, objectName: it.name });
      } else {
        // Rectangle déjà calculé à la composition (image entière, sans déformation).
        slideObj.addImage({ path: it.abs, x: round2(it.x), y: round2(it.y), w: round2(it.w), h: round2(it.h), altText: it.alt, objectName: it.name });
      }
      reg(idx, 'pic', [it.name]);
      if (it.radius) map[idx].radius = it.radius;
    } else if (it.k === 'text') {
      const runsOut = [];
      it.paras.forEach((p, pi) => {
        const ps = Math.round(paraSize(it, p, it.fitted) * 2) / 2;
        const F = faceOf(T, roleOf(it, p), paraWeight(it, p));
        const track = paraTrack(it, p, ps);
        const lnPt = Math.round(paraLineH(it, p, ps) * 72 * 100) / 100;
        const after = pi < it.paras.length - 1 ? Math.round(paraGap(it, p, ps) * 72 * 100) / 100 : 0;
        const runs = p.runs.length ? p.runs : [{ t: '' }];
        runs.forEach((r, ri) => {
          const o = {
            fontFace: xmlFace(F.face), fontSize: ps, bold: F.bold,
            color: r.accent ? c.P.accentTxt : (p.color || it.color),
            align: p.align || it.align, lineSpacing: lnPt, paraSpaceAfter: after, lang: langTag
          };
          if (track) o.charSpacing = Math.round(track * ps * 100) / 100;
          // [À FOURNIR] estompé : <a:alpha> dans le srgbClr du run (transparence 60 → alpha 40 000).
          // Ce remplissage porte un alpha : le dégradé du chiffre héros (post-traitement) le laisse.
          if (r.marker) o.transparency = Math.round((1 - STYLE.marqueur.texte) * 100);
          if (r.marker && r.ul !== false) o.underline = { style: STYLE.marqueur.trait, color: c.P.accentTxt };
          if (ri === runs.length - 1 && pi < it.paras.length - 1) o.breakLine = true;
          runsOut.push({ text: r.t, options: o });
        });
      });
      const F0 = faceOf(T, it.font, it.wt);
      slideObj.addText(runsOut, {
        x: round2(it.x), y: round2(it.y), w: round2(it.w), h: round2(it.h), margin: 0, valign: it.valign, align: it.align,
        isTextBox: true, objectName: it.name, wrap: !it.noWrap, fontFace: xmlFace(F0.face), fontSize: it.fitted, color: it.color, lang: langTag
      });
      reg(idx, 'sp', [it.name]);
      map[idx].text = true;
      map[idx].nparas = it.paras.length;
    }
  });
  return map;
}

// ---------------------------------------------------------------------------
// Rendu HTML (aperçu) : 1 pouce = 10 cqw, 1 pt = 1/7,2 cqw. Mêmes positions, corps, interlignes
// (line-height = interligne / corps), suivis et graisses que le .pptx.
// ---------------------------------------------------------------------------
function cq(inches) {
  return round2(inches * 10) + 'cqw';
}
function cqPt(pt) {
  return round2(pt / 7.2) + 'cqw';
}
function htmlItems(c, env) {
  const T = env.T;
  const out = [];
  c.items.forEach((it) => {
    const pos = 'left:' + cq(it.x !== undefined ? it.x : 0) + ';top:' + cq(it.y !== undefined ? it.y : 0) + ';width:' + cq(it.w || 0) + ';height:' + cq(it.h || 0) + ';';
    if (it.k === 'rect' || it.k === 'round' || it.k === 'ellipse') {
      let st = pos;
      if (it.fill) {
        const [r, g, b] = rgbOf(it.fill);
        st += 'background:rgba(' + r + ',' + g + ',' + b + ',' + (it.transparency ? round2(1 - it.transparency / 100) : 1) + ');';
      }
      if (it.line) st += 'border:' + cqPt(it.lineW || 1) + ' ' + (it.dash === 'dash' ? 'dashed' : 'solid') + ' #' + it.line + ';box-sizing:border-box;';
      if (it.k === 'ellipse') st += 'border-radius:50%;';
      if (it.k === 'round') st += 'border-radius:' + cq(Math.max(0, Math.min(it.radius || 0.1, it.w / 2, it.h / 2))) + ';';
      if (it.shadow) {
        const O = STYLE.schema.ombre;
        st += 'box-shadow:0 ' + cq(O.decalage) + ' ' + cq(O.flou) + ' rgba(0,0,0,' + O.opacite + ');';
      }
      out.push('<div class="sh" title="' + escHtml(it.name || '') + '" style="' + st + '"></div>');
    } else if (it.k === 'check') {
      const s = it.d;
      const box = 'left:' + cq(it.cx - s / 2) + ';top:' + cq(it.cy - s / 2) + ';width:' + cq(s) + ';height:' + cq(s) + ';';
      const pts = [[0.26, 0.5], [0.44, 0.68], [0.76, 0.3]].map((p) => (p[0] * 100) + ',' + (p[1] * 100)).join(' ');
      const sw = round2(Math.max(1.5, s * 7) / 72 / s * 100);
      out.push('<svg class="sh" style="' + box + '" viewBox="0 0 100 100"><polyline points="' + pts + '" fill="none" stroke="#' + it.color + '" stroke-width="' + sw + '" stroke-linecap="square"/></svg>');
    } else if (it.k === 'rule') {
      const [ur, ug, ub] = rgbOf(it.color);
      // Position calculée pour la police que l'aperçu affiche (xh, wh), pas pour celle du .pptx.
      const rx = it.xh !== undefined ? it.xh : it.x;
      const rw = it.wh !== undefined ? it.wh : it.w;
      out.push('<div class="sh" title="' + escHtml(it.name || '') + '" style="left:' + cq(rx) + ';top:' + cq(it.y) + ';width:' + cq(rw) + ';height:0;border-top:max(1px,' + cqPt(it.pt) + ') dotted rgba(' + ur + ',' + ug + ',' + ub + ',' + STYLE.marqueur.opacite + ');"></div>');
    } else if (it.k === 'image') {
      const data = 'data:' + it.info.mime + ';base64,' + it.info.buf.toString('base64');
      out.push('<img class="sh" alt="' + escHtml(it.alt || '') + '" src="' + data + '" style="' + pos + 'object-fit:' + (it.mode === 'cover' ? 'cover' : 'fill') + ';' + (it.radius ? 'border-radius:' + cq(it.radius) + ';' : '') + '">');
    } else if (it.k === 'text') {
      const jc = it.valign === 'top' ? 'flex-start' : it.valign === 'bottom' ? 'flex-end' : 'center';
      let inner = '';
      it.paras.forEach((p, pi) => {
        const ps = Math.round(paraSize(it, p, it.fitted) * 2) / 2;
        const role = roleOf(it, p);
        const F = faceOf(T, role, paraWeight(it, p));
        const fam = role === 'titre' ? T.cssTitre : T.cssTexte;
        const track = paraTrack(it, p, ps);
        const lineH = round2(paraLineH(it, p, ps) * 72 / ps);
        const after = pi < it.paras.length - 1 ? cq(paraGap(it, p, ps)) : '0';
        const st = 'font-size:' + cqPt(ps) + ';line-height:' + lineH + ';font-weight:' + F.css + ';font-family:' + escHtml(fam) +
          ';letter-spacing:' + (track ? track + 'em' : 'normal') +
          ';color:#' + (p.color || it.color) + ';text-align:' + (p.align || it.align) + ';margin:0 0 ' + after + ' 0;';
        const runs = p.runs.map((r) => {
          const t = escHtml(r.t) || '&#8203;';
          let s = '';
          if (r.marker) {
            // Estompé comme dans le .pptx : la couleur du texte remplacé, à l'opacité du marqueur. En
            // rgba plutôt qu'en opacity, pour que le pointillé garde la sienne, comme dans le .pptx.
            const [mr, mg, mb] = rgbOf(r.accent ? c.P.accentTxt : (p.color || it.color));
            s += 'color:rgba(' + mr + ',' + mg + ',' + mb + ',' + STYLE.marqueur.texte + ');';
          } else if (r.accent) s += 'color:#' + c.P.accentTxt + ';';
          if (r.grad && it.grad && !r.marker) s +='background:linear-gradient(90deg,#' + c.P.accent + ',#' + c.P.accent2 + ');-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;color:transparent;';
          if (r.marker && r.ul !== false) {
            const [ur, ug, ub] = rgbOf(c.P.accentTxt);
            s += 'text-decoration:underline dotted rgba(' + ur + ',' + ug + ',' + ub + ',' + STYLE.marqueur.opacite + ');text-decoration-thickness:max(2px,.05em);text-underline-offset:.12em;';
          }
          return s ? '<span' + (r.marker ? ' class="mk"' : '') + ' style="' + s + '">' + t + '</span>' : t;
        }).join('');
        inner += '<p style="' + st + (it.noWrap ? 'white-space:nowrap;' : '') + '">' + runs + '</p>';
      });
      out.push('<div class="tx" data-name="' + escHtml(it.name) + '" style="' + pos + 'justify-content:' + jc + ';"><div class="in">' + inner + '</div></div>');
    }
  });
  return out.join('\n');
}

function previewHtml(deck, composed, env, totals) {
  const meta = deck.meta || {};
  const L = LABELS[env.lang] || LABELS.fr;
  const cards = composed.map((x, i) => {
    const s = x.slide;
    const d = Number(s.duree_s) || 0;
    const head = (i + 1) + (s.bloc ? ' · ' + escHtml(s.bloc) : '') + ' · ' + fmtTime(x.start) + ' → ' + fmtTime(x.start + d) + ' <span class="id">' + escHtml(s.id || '') + '</span>';
    const warn = x.c.warnings.length ? '<ul class="warn">' + x.c.warnings.map((w) => '<li>' + escHtml(w) + '</li>').join('') + '</ul>' : '';
    const anim = x.steps && x.steps.length ? '<p class="anim">' + escHtml((env.lang === 'en' ? 'Builds in the .pptx: ' : 'Apparitions du .pptx : ') + describeSteps(x.steps, L)) + '</p>' : '';
    return '<section class="card" id="slide-' + (i + 1) + '">' +
      '<h2>' + head + '</h2>' +
      '<div class="slide" data-index="' + (i + 1) + '" style="background:#' + x.c.P.bg + '">' + htmlItems(x.c, env) + '</div>' +
      warn + anim + '<pre class="notes">' + escHtml(x.notes) + '</pre></section>';
  }).join('\n');
  const pubDesc = meta.public && typeof meta.public === 'object' ? (meta.public.description || meta.public.type) : meta.public;
  const target = Number(meta.duree_min) || 0;
  const info = [
    meta.format ? 'Format : ' + escHtml(meta.format) : '',
    pubDesc ? 'Public : ' + escHtml(pubDesc) + ' (' + escHtml(env.pub.cle) + ')' : '',
    meta.registre ? 'Registre : ' + escHtml(meta.registre) : '',
    'Durée : ' + fmtTime(totals.duree) + (target ? ' pour ' + target + ' min visées' : ''),
    composed.length + ' slides 16:9'
  ].filter(Boolean).join(' · ');
  return '<!doctype html>\n<html lang="' + (env.lang === 'en' ? 'en' : 'fr') + '">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n' +
    '<title>' + escHtml(meta.titre || 'Aperçu') + '</title>\n<style>\n' +
    ':root{--page:#ECEDEF;--ink:#1D1F22;--muted:#5F6670;--card:#FFFFFF;--warn:#8A4B00;--warnbg:#FFF4E0}\n' +
    '@media (prefers-color-scheme: dark){:root{--page:#16181B;--ink:#E9EBEE;--muted:#A3A9B1;--card:#212429;--warn:#FFC266;--warnbg:#3A2A10}}\n' +
    '*{box-sizing:border-box}html,body{margin:0}body{background:var(--page);color:var(--ink);font:15px/1.5 Arial,"Segoe UI",Roboto,sans-serif;padding:24px 16px 64px}\n' +
    'header{max-width:1100px;margin:0 auto 20px}header h1{font-size:22px;margin:0 0 4px}header p{margin:0;color:var(--muted)}\n' +
    '.card{max-width:1100px;margin:0 auto 28px;background:var(--card);border-radius:10px;padding:14px;box-shadow:0 1px 3px rgba(0,0,0,.08)}\n' +
    '.card h2{font-size:13px;font-weight:600;color:var(--muted);margin:0 0 10px;letter-spacing:.02em}.card h2 .id{font-weight:400;opacity:.7;margin-left:6px}\n' +
    '.slide{position:relative;width:100%;aspect-ratio:16/9;overflow:hidden;container-type:inline-size;border-radius:4px}\n' +
    '.slide .sh,.slide .tx{position:absolute;margin:0}\n.slide .tx{display:flex;flex-direction:column;overflow:visible}\n.slide .tx .in{width:100%}\n' +
    '.slide p{overflow-wrap:normal;word-break:normal;hyphens:none;font-kerning:normal}\n' +
    '.notes{white-space:pre-wrap;font:13px/1.55 ui-monospace,Menlo,Consolas,monospace;margin:12px 0 0;padding:12px;background:var(--page);border-radius:6px;color:var(--ink)}\n' +
    '.warn{margin:10px 0 0;padding:8px 12px 8px 28px;background:var(--warnbg);color:var(--warn);border-radius:6px;font-size:13px}\n' +
    '.anim{margin:10px 0 0;font-size:13px;color:var(--muted)}\n' +
    'body.qa .tx.over{outline:2px solid #FF2D2D;outline-offset:1px}\n#qa-banner{display:none;position:sticky;top:0;z-index:5;max-width:1100px;margin:0 auto 16px;padding:8px 12px;border-radius:6px;background:#FF2D2D;color:#fff;font-weight:600}\nbody.qa #qa-banner{display:block}\n' +
    '@media print{body{padding:0;background:#fff}.card{box-shadow:none;page-break-inside:avoid}}\n' +
    '</style>\n</head>\n<body>\n<div id="qa-banner"></div>\n<header><h1>' + escHtml(meta.titre || 'Aperçu') + '</h1><p>' + info + '</p></header>\n' +
    cards + '\n<script>\n' +
    '(function(){\n' +
    '  function overflows(){var out=[];document.querySelectorAll(".slide").forEach(function(sl){sl.querySelectorAll(".tx").forEach(function(b){var i=b.firstElementChild;var bh=b.clientHeight,bw=b.clientWidth;var over=i.offsetHeight>bh+1||i.scrollWidth>bw+1;b.classList.toggle("over",over);if(over)out.push({slide:+sl.dataset.index,element:b.dataset.name,hauteur:i.offsetHeight,place:bh,largeur:i.scrollWidth,placeLargeur:bw});});});return out;}\n' +
    '  window.__deckQA=overflows;\n' +
    '  function run(){var o=overflows();if(location.hash.indexOf("qa")>=0){document.body.classList.add("qa");var bn=document.getElementById("qa-banner");bn.style.background=o.length?"#D92D20":"#1F7A4D";bn.textContent=o.length?o.length+" texte(s) débordent : slides "+o.map(function(x){return x.slide}).join(", "):"Aucun débordement détecté.";}}\n' +
    '  if(document.fonts&&document.fonts.ready){document.fonts.ready.then(run);}else{window.addEventListener("load",run);}\n' +
    '  window.addEventListener("hashchange",run);window.addEventListener("resize",run);\n' +
    '})();\n</script>\n</body>\n</html>\n';
}

// ---------------------------------------------------------------------------
// Post-traitement du paquet : couleurs du thème, notes en paragraphes, placeholders groupés,
// dégradé du chiffre héros, transition en fondu et animations d'entrée de chaque slide.
// Seuls les fichiers des slides, du thème et des notes sont réécrits ; presentation.xml
// n'est pas touché.
// ---------------------------------------------------------------------------
const EMU = 914400;
function slideTokens(xml) {
  const toks = [];
  const re = /<p:(sp|pic|cxnSp|graphicFrame|grpSp)\b[^>]*>[\s\S]*?<\/p:\1>/g;
  let m;
  while ((m = re.exec(xml))) {
    const idm = /<p:cNvPr id="(\d+)" name="([^"]*)"/.exec(m[0]);
    toks.push({ start: m.index, end: m.index + m[0].length, tag: m[1], id: idm ? Number(idm[1]) : -1, name: idm ? unescXml(idm[2]) : '', xml: m[0] });
  }
  return toks;
}
function effectXml(t, nodeType, delay, nid, dur) {
  const effId = nid();
  const setId = nid();
  const animId = nid();
  const tgt = t.pr
    ? '<p:spTgt spid="' + t.spid + '"><p:txEl><p:pRg st="' + t.pr[0] + '" end="' + t.pr[1] + '"/></p:txEl></p:spTgt>'
    : '<p:spTgt spid="' + t.spid + '"/>';
  const grp = t.type === 'sp' ? ' grpId="0"' : '';
  return '<p:par><p:cTn id="' + effId + '" presetID="10" presetClass="entr" presetSubtype="0" fill="hold"' + grp + ' nodeType="' + nodeType + '">' +
    '<p:stCondLst><p:cond delay="' + delay + '"/></p:stCondLst><p:childTnLst>' +
    '<p:set><p:cBhvr><p:cTn id="' + setId + '" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>' +
    '<p:tgtEl>' + tgt + '</p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>' +
    '<p:to><p:strVal val="visible"/></p:to></p:set>' +
    '<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="' + animId + '" dur="' + dur + '"/><p:tgtEl>' + tgt + '</p:tgtEl></p:cBhvr></p:animEffect>' +
    '</p:childTnLst></p:cTn></p:par>';
}
function timingXml(steps) {
  const DUR = STYLE.apparition.duree_ms;
  const SUITE = STYLE.apparition.suite_ms;
  const clicks = [];
  steps.forEach((s) => {
    if (!s.targets.length) return;
    if (s.trigger === 'click' || !clicks.length) clicks.push([{ trig: 'click', steps: [s] }]);
    else {
      const subs = clicks[clicks.length - 1];
      if (s.trigger === 'after') subs.push({ trig: 'after', steps: [s] });
      else subs[subs.length - 1].steps.push(s);
    }
  });
  if (!clicks.length) return { xml: '', builds: [] };
  let id = 2;
  const nid = () => ++id;
  const builds = [];
  let body = '';
  clicks.forEach((subs) => {
    const outerId = nid();
    let inner = '';
    let offset = 0;
    subs.forEach((sg, k) => {
      const effDelay = sg.trig === 'after' ? SUITE : 0;
      const subOffset = k === 0 ? 0 : offset;
      const subId = nid();
      let effects = '';
      let first = true;
      sg.steps.forEach((st) => st.targets.forEach((t) => {
        const nodeType = first ? (sg.trig === 'after' ? 'afterEffect' : 'clickEffect') : 'withEffect';
        first = false;
        effects += effectXml(t, nodeType, effDelay, nid, DUR);
        builds.push(t);
      }));
      inner += '<p:par><p:cTn id="' + subId + '" fill="hold"><p:stCondLst><p:cond delay="' + subOffset + '"/></p:stCondLst><p:childTnLst>' + effects + '</p:childTnLst></p:cTn></p:par>';
      offset = subOffset + effDelay + DUR;
    });
    body += '<p:par><p:cTn id="' + outerId + '" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>' + inner + '</p:childTnLst></p:cTn></p:par>';
  });
  const xml = '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
    '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' + body + '</p:childTnLst></p:cTn>' +
    '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
    '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>' +
    '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>';
  return { xml, builds };
}
// pptxgenjs répète <a:pPr> devant chaque run d'un paragraphe (texte en accent, point final) :
// le schéma n'en admet qu'un, en tête. On garde le premier, on retire les suivants.
function fixParagraphs(xml) {
  const PPR = /<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/g;
  return xml.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (all, inner) => {
    let seen = false;
    const out = inner.replace(PPR, (m, off) => {
      if (!seen && !/<a:(?:r|br|fld)\b/.test(inner.slice(0, off))) {
        seen = true;
        return m;
      }
      return '';
    });
    return '<a:p>' + out + '</a:p>';
  });
}
function processSlideXml(xml, x, env, slideNo) {
  xml = fixParagraphs(xml);
  // pptxgenjs coupe le crénage (kern="0") dès qu'un suivi est posé : un titre serré sans crénage
  // laisse des trous entre certaines paires (To, Av, Te). On le rétablit dès 12 pt.
  xml = xml.replace(/ kern="0"/g, ' kern="1200"');
  const c = x.c;
  const map = x.map || [];
  const warn = (m) => env.globalWarnings.push('slide ' + slideNo + ' : ' + m);
  let toks = slideTokens(xml);
  const byId = {};
  toks.forEach((t) => { byId[t.id] = t; });
  // 1. Les identifiants calculés au rendu correspondent-ils aux formes écrites ?
  let consistent = true;
  map.forEach((m) => {
    if (!m) return;
    m.ids.forEach((id, k) => {
      const t = byId[id];
      if (!t || t.name !== m.names[k]) consistent = false;
    });
  });
  if (!consistent) warn('identifiants des formes inattendus : placeholders non groupés et apparitions non posées.');
  // 2. Placeholders dessinés : leurs formes deviennent un seul groupe, simple à remplacer.
  const phGroup = {};
  if (consistent && c.phs.length) {
    let nextId = Math.max.apply(null, toks.map((t) => t.id).concat([1])) + 1;
    const groups = [];
    c.phs.forEach((ph) => {
      const idxs = c.items.map((it, i) => (it.ph === ph.id ? i : -1)).filter((i) => i >= 0);
      const ids = [].concat.apply([], idxs.map((i) => map[i].ids));
      if (ids.length < 2) return;
      const tks = ids.map((id) => byId[id]).sort((a, b) => a.start - b.start);
      const contiguous = tks.every((t, k) => k === 0 || toks.indexOf(t) === toks.indexOf(tks[k - 1]) + 1);
      if (!contiguous) return;
      const its = idxs.map((i) => c.items[i]);
      const bx = (it) => (it.k === 'check' ? [it.cx - it.d / 2, it.cy - it.d / 2, it.d, it.d] : [it.x, it.y, it.w, it.h]);
      const x0 = Math.min.apply(null, its.map((it) => bx(it)[0]));
      const y0 = Math.min.apply(null, its.map((it) => bx(it)[1]));
      const x1 = Math.max.apply(null, its.map((it) => bx(it)[0] + bx(it)[2]));
      const y1 = Math.max.apply(null, its.map((it) => bx(it)[1] + bx(it)[3]));
      const gid = nextId++;
      groups.push({ gid, name: ph.name, first: tks[0], last: tks[tks.length - 1], box: [x0, y0, Math.max(0.01, x1 - x0), Math.max(0.01, y1 - y0)] });
      phGroup[ph.id] = gid;
    });
    groups.sort((a, b) => b.first.start - a.first.start).forEach((g) => {
      const [gx, gy, gw, gh] = g.box.map((v) => Math.round(v * EMU));
      const xf = '<a:xfrm><a:off x="' + gx + '" y="' + gy + '"/><a:ext cx="' + Math.max(1, gw) + '" cy="' + Math.max(1, gh) + '"/><a:chOff x="' + gx + '" y="' + gy + '"/><a:chExt cx="' + Math.max(1, gw) + '" cy="' + Math.max(1, gh) + '"/></a:xfrm>';
      const open = '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="' + g.gid + '" name="' + escXml(g.name) + '"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>' + xf + '</p:grpSpPr>';
      xml = xml.slice(0, g.first.start) + open + xml.slice(g.first.start, g.last.end) + '</p:grpSp>' + xml.slice(g.last.end);
    });
    toks = slideTokens(xml);
    toks.forEach((t) => { byId[t.id] = t; });
  }
  // 3. Dégradé du chiffre héros (accent → accent2), dans les propriétés du texte.
  if (consistent) {
    c.items.forEach((it, i) => {
      if (!it.grad || !map[i]) return;
      const id = map[i].ids[0];
      const re = new RegExp('<p:sp>(?:(?!<\\/p:sp>)[\\s\\S])*?<p:cNvPr id="' + id + '" [\\s\\S]*?<\\/p:sp>');
      xml = xml.replace(re, (sp) => sp.replace(new RegExp('(<a:rPr\\b[^>]*>)<a:solidFill><a:srgbClr val="' + c.P.accentTxt + '"(?:\\/>|><\\/a:srgbClr>)<\\/a:solidFill>', 'g'),
        '$1<a:gradFill flip="none" rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="' + c.P.accent + '"/></a:gs><a:gs pos="100000"><a:srgbClr val="' + c.P.accent2 + '"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>'));
    });
  }
  // 3 bis. Coins arrondis d'une capture nue (pptxgenjs n'écrit qu'un rectangle).
  if (consistent) {
    c.items.forEach((it, i) => {
      if (it.k !== 'image' || !it.radius || !map[i]) return;
      const id = map[i].ids[0];
      const adj = Math.round(Math.min(50000, it.radius / Math.max(0.01, Math.min(it.w, it.h)) * 100000));
      const re = new RegExp('<p:pic>(?:(?!<\\/p:pic>)[\\s\\S])*?<p:cNvPr id="' + id + '" [\\s\\S]*?<\\/p:pic>');
      xml = xml.replace(re, (pic) => pic.replace(/<a:prstGeom prst="rect">\s*<a:avLst\s*\/>\s*<\/a:prstGeom>/, '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ' + adj + '"/></a:avLst></a:prstGeom>'));
    });
  }
  // 3 ter. Soulignement des [À FOURNIR] adouci (opacité de STYLE.marqueur).
  const alpha = Math.round(STYLE.marqueur.opacite * 100000);
  if (alpha < 100000) xml = xml.replace(/<a:uFill><a:solidFill><a:srgbClr val="([0-9A-F]{6})"(?:\/>|><\/a:srgbClr>)<\/a:solidFill><\/a:uFill>/g, '<a:uFill><a:solidFill><a:srgbClr val="$1"><a:alpha val="' + alpha + '"/></a:srgbClr></a:solidFill></a:uFill>');
  // 4. Apparitions : cibles → identifiants (le groupe pour un placeholder), plages de paragraphes vérifiées.
  let timing = '';
  if (consistent && x.steps && x.steps.length) {
    const npara = (id) => {
      const t = byId[id];
      if (!t) return -1;
      const body = /<p:txBody>[\s\S]*<\/p:txBody>/.exec(t.xml);
      return body ? (body[0].match(/<a:p[ >]/g) || []).length : 0;
    };
    const resolved = x.steps.map((s) => {
      const seen = new Set();
      const targets = [];
      s.targets.forEach((tg) => {
        const it = c.items[tg.idx];
        const m = map[tg.idx];
        if (!it || !m) return;
        let list;
        if (it.ph !== undefined && phGroup[it.ph]) list = [{ spid: phGroup[it.ph], type: 'grp', pr: null }];
        else list = m.ids.map((id) => ({ spid: id, type: m.type, pr: tg.pr }));
        list.forEach((t) => {
          if (t.pr && npara(t.spid) !== (m.nparas || 0)) t.pr = null;
          const key = t.spid + ':' + (t.pr ? t.pr.join('-') : '*');
          if (seen.has(key)) return;
          seen.add(key);
          targets.push(t);
        });
      });
      return { trigger: s.trigger, targets };
    });
    // Une forme animée en entier et par paragraphes à la fois : on garde l'animation entière.
    const wholeIds = new Set();
    resolved.forEach((s) => s.targets.forEach((t) => { if (!t.pr) wholeIds.add(t.spid); }));
    const seenWhole = new Set();
    resolved.forEach((s) => {
      s.targets = s.targets.filter((t) => {
        if (t.pr && wholeIds.has(t.spid)) return false;
        if (!t.pr) {
          if (seenWhole.has(t.spid)) return false;
          seenWhole.add(t.spid);
        }
        return true;
      });
    });
    const tm = timingXml(resolved);
    if (tm.xml) {
      const bld = {};
      tm.builds.forEach((t) => {
        if (t.type !== 'sp') return;
        const isText = map.some((m) => m && m.text && m.ids.includes(t.spid));
        const cur = bld[t.spid] || (isText ? '' : ' animBg="1"');
        bld[t.spid] = t.pr ? ' build="p"' : cur;
      });
      const entries = Object.keys(bld).map((id) => '<p:bldP spid="' + id + '" grpId="0"' + bld[id] + '/>').join('');
      timing = tm.xml + (entries ? '<p:bldLst>' + entries + '</p:bldLst>' : '') + '</p:timing>';
      x.clicks = resolved.filter((s) => s.trigger === 'click' && s.targets.length).length;
    }
  }
  // 5. Transition en fondu, puis animations : après <p:clrMapOvr>, avant </p:sld>.
  const transition = '<p:transition spd="' + STYLE.transition.vitesse + '"><p:fade/></p:transition>';
  const tail = '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>';
  if (xml.includes(tail + '</p:sld>')) xml = xml.replace(tail + '</p:sld>', tail + transition + timing + '</p:sld>');
  else {
    xml = xml.replace(/<\/p:sld>\s*$/, transition + timing + '</p:sld>');
    if (!xml.includes('<p:transition')) warn('transition non posée (structure de slide inattendue).');
  }
  return xml;
}
async function postProcess(buf, env, composed) {
  const JSZip = loadJSZip();
  if (!JSZip) {
    env.globalWarnings.push('jszip introuvable : couleurs du thème, notes, transitions et apparitions non posées (le fichier reste valide).');
    return buf;
  }
  const zip = await JSZip.loadAsync(buf);
  const T = env.T;
  const themePath = 'ppt/theme/theme1.xml';
  const tf = zip.file(themePath);
  if (tf) {
    let xml = await tf.async('string');
    const slots = {
      dk1: T.texte, lt1: T.fondClair, dk2: T.fondSombre, lt2: mix(T.fondClair, T.texte, 0.08),
      accent1: T.accent, accent2: T.accent2, accent3: T.texteSec, accent4: T.texteSecSombre,
      accent5: mix(T.accent, T.fondSombre, 0.4), accent6: mix(T.accent2, T.fondSombre, 0.4), hlink: T.accent, folHlink: T.accent2
    };
    const inner = Object.keys(slots).map((k) => '<a:' + k + '><a:srgbClr val="' + slots[k] + '"/></a:' + k + '>').join('');
    const safeName = escXml(String(T.nom).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, ''));
    xml = xml.replace(/<a:clrScheme name="[^"]*">[\s\S]*?<\/a:clrScheme>/, '<a:clrScheme name="' + safeName + '">' + inner + '</a:clrScheme>');
    xml = xml.replace(/(<a:theme\b[^>]*\bname=")[^"]*(")/, '$1' + safeName + '$2');
    zip.file(themePath, xml);
  }
  const langTag = env.lang === 'en' ? 'en-US' : 'fr-FR';
  const notesFiles = Object.keys(zip.files).filter((n) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n));
  for (const n of notesFiles) {
    let xml = await zip.file(n).async('string');
    xml = xml.replace(/<a:p><a:r><a:rPr lang="en-US" dirty="0"\/><a:t>([\s\S]*?)<\/a:t><\/a:r><a:endParaRPr lang="en-US" dirty="0"\/><\/a:p>/, (all, text) => {
      return text.split(/\r?\n/).map((line) => {
        if (!line) return '<a:p><a:endParaRPr lang="' + langTag + '" dirty="0"/></a:p>';
        const sp = /^\s|\s$/.test(line) ? ' xml:space="preserve"' : '';
        return '<a:p><a:r><a:rPr lang="' + langTag + '" dirty="0"/><a:t' + sp + '>' + line + '</a:t></a:r></a:p>';
      }).join('');
    });
    zip.file(n, xml);
  }
  for (let i = 0; i < composed.length; i++) {
    const file = 'ppt/slides/slide' + (i + 1) + '.xml';
    const f = zip.file(file);
    if (!f) continue;
    const xml = await f.async('string');
    zip.file(file, processSlideXml(xml, composed[i], env, i + 1));
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// ---------------------------------------------------------------------------
// Programme principal
// ---------------------------------------------------------------------------
function usage(code) {
  console.log('Usage : node build_deck.js deck.json sortie.pptx [--preview sortie.html] [--polices auto|repli]');
  console.log('  --polices auto (par défaut) : le .pptx écrit les polices du thème, installées ici ou non ; repli : il écrit les polices de repli.');
  process.exit(code);
}

async function main() {
  const args = process.argv.slice(2);
  if (!args.length || args.includes('-h') || args.includes('--help')) usage(args.length ? 0 : 2);
  const pos = [];
  let preview = null;
  let fontMode = 'auto';
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--preview') preview = args[++i];
    else if (args[i] === '--polices') fontMode = args[++i];
    else pos.push(args[i]);
  }
  if (pos.length < 2) usage(2);
  // « principale » (ancien mode qui forçait la police du thème) vaut désormais « auto ».
  if (fontMode === 'principale') fontMode = 'auto';
  if (!['auto', 'repli'].includes(fontMode)) {
    console.error('--polices attend auto ou repli.');
    process.exit(2);
  }
  const deckPath = path.resolve(pos[0]);
  const outPath = path.resolve(pos[1]);
  let deck;
  try {
    deck = JSON.parse(fs.readFileSync(deckPath, 'utf8'));
  } catch (e) {
    console.error('Impossible de lire ' + deckPath + ' : ' + e.message);
    process.exit(2);
  }
  if (!Array.isArray(deck.slides) || !deck.slides.length) {
    console.error('deck.json ne contient aucune slide (champ « slides »).');
    process.exit(2);
  }
  const meta = deck.meta || {};
  const globalWarnings = [];
  const fontNotes = [];
  const lang = String(meta.langue || 'fr').toLowerCase().startsWith('en') ? 'en' : 'fr';
  if (meta.ratio && String(meta.ratio).replace(/\s/g, '') !== '16:9') globalWarnings.push('meta.ratio vaut « ' + meta.ratio + ' » : le rendu est toujours en 16:9.');
  const T = resolveTheme(deck.theme, fontMode, globalWarnings, fontNotes);
  FONT.titre = { cle: T.cleTitre, mesure: T.mesureTitre, ici: T.iciTitre };
  FONT.texte = { cle: T.cleTexte, mesure: T.mesureTexte, ici: T.iciTexte };
  const pub = resolvePublic(meta, loadPublics());
  let sourceMode = meta.affichage_source || pub.affichage_source || 'sous_le_chiffre';
  if (!['sous_le_chiffre', 'renvoi', 'carton_final', 'aucun'].includes(sourceMode)) sourceMode = 'sous_le_chiffre';
  const intervenants = arr(deck.intervenants);
  const people = {};
  intervenants.forEach((p) => { if (p && p.id) people[p.id] = p; });

  const env = {
    T, lang, deckDir: path.dirname(deckPath), pub, sourceMode, intervenants, people,
    collectedSources: [], sourceCounter: 0, sourceIndex: {}, lastSpeaker: null, introduced: new Set(), globalWarnings,
    finalIndex: deck.slides.length - 1
  };

  // Composition
  const composed = [];
  let clock = 0;
  deck.slides.forEach((slide, i) => {
    const c = composeSlide(deck, slide, i, env);
    if (c.sourceRef && c.sourceNew) env.collectedSources.push({ n: c.sourceRef, text: c.sourceText, slides: [i + 1] });
    else if (c.sourceRef) env.collectedSources.find((x) => x.n === c.sourceRef).slides.push(i + 1);
    const steps = planAnimations(c);
    const notes = buildNotes(slide, c, env, clock);
    composed.push({ slide, c, notes, steps, start: clock });
    clock += Number(slide.duree_s) || 0;
  });
  if ((sourceMode === 'carton_final' || sourceMode === 'renvoi') && env.collectedSources.length && meta.carton_sources !== false) {
    const slide = { id: 'sources', bloc: composed[composed.length - 1].slide.bloc, archetype: 'SOURCES', fond: 'sombre', duree_s: 0 };
    const c = new SlideCtx(deck, slide, composed.length, env);
    c.sourceOnSlide = false;
    R.SOURCES(c);
    composed.push({ slide, c, notes: LABELS[lang].finalNotes, steps: [], start: clock });
  }

  // .pptx
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';
  pres.title = stripValider(str(meta.titre || 'Présentation'));
  if (intervenants[0] && intervenants[0].nom) pres.author = String(intervenants[0].nom);
  pres.theme = { headFontFace: xmlFace(T.fontTitre), bodyFontFace: xmlFace(T.fontTexte) };
  pres.defineSlideMaster({ title: 'Fond sombre', background: { color: T.fondSombre } });
  pres.defineSlideMaster({ title: 'Fond clair', background: { color: T.fondClair } });
  let lastSection = null;
  composed.forEach((x) => {
    const bloc = str(x.slide.bloc).trim();
    if (bloc && bloc !== lastSection) {
      pres.addSection({ title: bloc });
      lastSection = bloc;
    }
    const opts = { masterName: x.c.P.fond === 'clair' ? 'Fond clair' : 'Fond sombre' };
    if (lastSection) opts.sectionTitle = lastSection;
    const so = pres.addSlide(opts);
    so.background = { color: x.c.P.bg };
    x.map = pptxRender(pres, so, x.c, env);
    so.addNotes(x.notes);
  });
  let buf = await pres.write({ outputType: 'nodebuffer' });
  buf = await postProcess(buf, env, composed);
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, buf);

  // Aperçu
  if (preview) {
    const prevPath = path.resolve(preview);
    fs.mkdirSync(path.dirname(prevPath), { recursive: true });
    fs.writeFileSync(prevPath, previewHtml(deck, composed, env, { duree: clock }));
  }

  // Compte rendu
  const L = LABELS[lang];
  const target = Number(meta.duree_min) || 0;
  const animated = composed.filter((x) => x.clicks);
  console.log('Deck rendu : ' + composed.length + ' slides 16:9 → ' + outPath + ' (un seul fichier)');
  console.log('Transitions : fondu sur chaque slide · apparitions : ' + animated.reduce((n, x) => n + x.clicks, 0) + ' clic(s) sur ' + animated.length + ' slide(s)');
  if (preview) console.log('Aperçu HTML → ' + path.resolve(preview) + ' (ajoutez #qa à l\'adresse pour voir les débordements)');
  const semiNote = (fam, semi) => '« ' + fam + ' »' + (semi ? ' (demi-gras : « ' + semi + ' »)' : ' (demi-gras rendu en gras : pas de face SemiBold installée)');
  console.log('Polices : titre ' + semiNote(T.fontTitre, T.semiTitre) + ', texte ' + semiNote(T.fontTexte, T.semiTexte) + ' · public : ' + pub.cle + ' · sources : ' + sourceMode);
  fontNotes.forEach((m) => console.log('  ' + m));
  console.log('Minutage : ' + fmtTime(clock) + (target ? ' pour ' + target + ' min visées' : ''));
  // Ce qu'il faut valider et fournir, slide par slide.
  const val = [];
  const prov = [];
  composed.forEach((x, i) => {
    const notesProps = [];
    noteSegments(x.slide, env).forEach((s) => stripValider(s.texte, notesProps));
    const screen = x.c.proposals.map((p) => L.q(p));
    const said = notesProps.filter((p) => !x.c.proposals.includes(p)).map((p) => L.q(p));
    if (screen.length || said.length) val.push('  - slide ' + (i + 1) + ' (' + (x.slide.id || '') + ') : ' + (screen.length ? (lang === 'en' ? 'on screen ' : 'à l\'écran ') + screen.join(' ; ') : '') + (screen.length && said.length ? ' · ' : '') + (said.length ? (lang === 'en' ? 'in the script ' : 'dans le script ') + said.join(' ; ') : ''));
    const items = provideItems(x.slide, x.c, env);
    const inNotes = noteSegments(x.slide, env).reduce((n, s) => n + fournirLabels(s.texte).length, 0);
    if (items.length || inNotes) prov.push('  - slide ' + (i + 1) + ' (' + (x.slide.id || '') + ') : ' + items.join(' ') + (inNotes ? (items.length ? ' · ' : '') + inNotes + (lang === 'en' ? ' marker(s) in the script' : ' marqueur(s) dans le script') : ''));
  });
  console.log(L.toValidate + ' (' + val.length + ' slide' + (val.length > 1 ? 's' : '') + ')' + (val.length ? ' :' : ' : aucun'));
  val.forEach((l) => console.log(l));
  console.log(L.toProvide + ' (' + prov.length + ' slide' + (prov.length > 1 ? 's' : '') + ')' + (prov.length ? ' :' : ' : aucun'));
  prov.forEach((l) => console.log(l));
  const all = [];
  globalWarnings.forEach((w) => all.push('  - ' + w));
  composed.forEach((x, i) => x.c.warnings.forEach((w) => all.push('  - slide ' + (i + 1) + ' (' + (x.slide.id || '') + ') : ' + w)));
  composed.forEach((x, i) => x.c.fournirs.forEach((l) => {
    const n = l.split(/\s+/).filter(Boolean).length;
    if (n > 3) all.push('  - slide ' + (i + 1) + ' (' + (x.slide.id || '') + ') : libellé [À FOURNIR] de ' + n + ' mots à l\'écran (« ' + l + ' ») : 3 au plus, la consigne va dans a_fournir.');
  }));
  if (all.length) {
    console.log('À vérifier (' + all.length + ') :');
    all.forEach((l) => console.log(l));
  }
}

main().catch((e) => {
  console.error('Échec du rendu : ' + (e && e.stack ? e.stack : e));
  process.exit(1);
});
