# sebmakes-keynote-likeapple-pls

Une skill Claude qui construit une présentation 16:9 **et le script à dire**, selon la grammaire des keynotes Apple : un déroulé qui monte en puissance, une séquence d'annonce, une idée par écran, des relais entre intervenants, des récapitulatifs emboîtés. Elle reprend la grammaire, jamais la marque : les couleurs, les polices, le produit et les chiffres sont les vôtres.

*English below.*

## Ce qu'elle fait

1. **Elle vous pose les bonnes questions** : le format (Keynote, PowerPoint, Figma Slides, Canva ou vidéo), l'habillage (un lien vers votre design system, ou vos couleurs en HEX), la typologie de public, le tutoiement ou le vouvoiement, puis le brief.
2. **Elle vous montre le conducteur** (blocs, intervenants, minutage) avant de produire quoi que ce soit.
3. **Elle produit le deck** : un .pptx 16:9 avec fondus et apparitions au clic, un aperçu HTML, et le script complet dans les notes et dans `script.md`.

Quelques règles qu'elle tient :

- **Aucun chiffre inventé.** Un chiffre, un prix, une date, un client ou un témoignage qui manque reste visible comme `[À FOURNIR]`, estompé à l'écran.
- **Des propositions signalées.** Quand le brief est mince, elle propose des formulations marquées `[À VALIDER]`.
- **Le public décide des procédés.** Graphiques sans graduation, chiffres sans source, concurrents anonymes : ce qui passe devant le grand public ne passe pas devant des investisseurs ou des acheteurs. La skill règle tout cela d'après la typologie de public.

La grammaire vient d'une analyse de huit keynotes Apple publiques (2019-2026) : déroulés, plans, formules en anglais et en français. Les mises en page sont étalonnées sur des mesures prises sur ces plans. La skill ne contient aucune image, aucune police ni aucune citation d'Apple.

## Installation

**Claude Code, comme plugin :**

```bash
claude plugin marketplace add sebmakes/sebmakes-keynote-likeapple-pls
claude plugin install sebmakes-keynote-likeapple-pls@sebmakes
```

**Claude Code, à la main :** copiez le dossier `skills/sebmakes-keynote-likeapple-pls` dans `~/.claude/skills/`.

**Claude.ai :** téléchargez le fichier `.skill` de la dernière version (onglet *Releases*), puis importez-le dans vos réglages de skills.

**Prérequis :** Node 18 ou plus (la bibliothèque pptxgenjs s'installe toute seule au premier rendu), Python 3.9 ou plus pour le contrôle du deck.

## Utilisation

Demandez simplement, par exemple :

> Fais-moi une présentation à la Apple de 10 minutes pour présenter notre nouveau module de facturation à nos clients. Couleurs : #091230 et #E4FF3A. On vouvoie le public. En PowerPoint.

Ce qui reste à vérifier à la main : les apparitions dans Keynote ou Canva (elles ne passent pas toujours à l'import), et les polices, à installer sur la machine qui présente.

---

## English

A Claude skill that builds a 16:9 deck **and the full speaker script** with the grammar of Apple keynotes: a run of show that builds up, an announcement sequence, one idea per slide, handovers between speakers, nested recaps. It takes the grammar, never the brand: colours, fonts, product and figures are yours.

- It asks for the format (Keynote, PowerPoint, Figma Slides, Canva or video), your branding (a design-system link or HEX colours), the audience type, and the tone. Then it shows the run of show before producing anything.
- It delivers a 16:9 .pptx with fades and click builds, an HTML preview, and the full script in the speaker notes and in `script.md`.
- It never invents a figure, price, date, client or testimonial: missing items stay visible as `[TO PROVIDE]`. Proposed wording is flagged `[TO VALIDATE]`.

Install with `claude plugin marketplace add sebmakes/sebmakes-keynote-likeapple-pls`, then `claude plugin install sebmakes-keynote-likeapple-pls@sebmakes`. You can also copy `skills/sebmakes-keynote-likeapple-pls` into `~/.claude/skills/`. Requires Node 18+ and Python 3.9+.

---

Not affiliated with, endorsed or sponsored by Apple Inc. Apple and Keynote are trademarks of Apple Inc., registered in the U.S. and other countries. Licence MIT.
