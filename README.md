# sebmake-keynote-likeapple-pls

A Claude skill that builds a 16:9 deck the way Apple builds its keynotes, then writes the script you'll read over it. The structure comes from eight public Apple keynotes (2019 to 2026): a run of show that saves the biggest news for last, announcements that unfold in set steps, one idea per slide, speakers handing off to each other, recaps nested inside recaps. Apple's look stays out of it; your deck gets your colors, fonts, product and numbers.

The skill's own instructions are in French, but it builds decks in English or French.

## What it does

It opens with a few questions you can answer in one go: output format (Keynote, PowerPoint, Figma Slides, Canva or video), branding as a design-system link or HEX colors, audience type, and tone, which means first names or formal titles in English and *tu* or *vous* in French. A short brief follows, and you can leave parts of it blank.

Before rendering anything, it shows the run of show as a table of blocks, speakers, timing and slide count, so you can reorder it while that's still cheap.

You get a 16:9 .pptx with fades and click builds, an HTML preview, and the full script twice: in the speaker notes and in `script.md`.

## What it won't do

Make up a number. A missing figure, price, date, customer name or testimonial stays on the slide as `[TO PROVIDE: …]`, dimmed so it reads as an empty slot (`[À FOURNIR]` in a French deck). Wording gets different treatment: when your brief is thin, the skill drafts the thesis or a feature description itself and flags it `[TO VALIDATE: …]` until you sign off.

It also won't pull the same tricks on every audience. Charts with no axis labels, unsourced figures, an unnamed competitor, "up to" claims all play fine at a consumer launch and get picked apart by investors or a procurement team. You choose the audience type once, and the skill decides what each slide can get away with.

Type sizes, weights, letter spacing and positions come from 67 frames of those keynotes, measured pixel by pixel. The repo holds no Apple images, fonts or quotes.

## Layouts

25 slide types taken from the keynotes (title card, hero number, bento recap, product alone, screenshot inside a device, full-bleed photo, pricing, availability and more) share 147 layout variants, and each type has its own entry in `references/designs.md`. The skill picks a layout for every slide and varies it from one block to the next.

It also hands you a list of visuals to prepare, each with a brief: type, framing, transparent background or not, resolution. Until a visual shows up, the slide carries a neutral placeholder with no numbers and no logo.

![137 of the 147 layouts, rendered by the skill for a fictional brand](skills/sebmake-keynote-likeapple-pls/assets/galerie.jpg)

## Install

Claude Code, as a plugin:

```bash
claude plugin marketplace add sebmake/sebmake-keynote-likeapple-pls
claude plugin install sebmake-keynote-likeapple-pls@sebmake
```

Claude Code, by hand: copy the `skills/sebmake-keynote-likeapple-pls` folder into `~/.claude/skills/`.

Claude.ai: download the `.skill` file from the latest release, then upload it in your skill settings.

You'll need Node 18 or later and Python 3.9 or later. The skill runs `npm install` for pptxgenjs on the first render, and Python runs the deck check.

## Usage

Ask the way you'd ask a colleague:

> Make me a 10-minute Apple-style presentation introducing our new building-maintenance service to property managers. Colors: #0F1B2A and #1C9461. Formal tone. PowerPoint.

Two things still need a human eye. Click builds don't always survive an import into Keynote or Canva, and the fonts have to be installed on whatever machine you present from.

---

Not affiliated with, endorsed or sponsored by Apple Inc. Apple and Keynote are trademarks of Apple Inc., registered in the U.S. and other countries. MIT License.
