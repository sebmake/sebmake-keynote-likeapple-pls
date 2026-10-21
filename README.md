# sebmakes-keynote-likeapple-pls

A Claude skill that builds a 16:9 deck the way Apple builds its keynotes, then writes the script you'll read over it. The structure comes from eight public Apple keynotes (2019 to 2026): a run of show that saves the biggest news for last, an announcement that unfolds in set steps, one idea per slide, speakers handing off to each other, and recaps nested inside recaps. None of Apple's look comes along. Your deck gets your colors, your fonts, your product and your numbers.

The skill's own instructions are in French, but it builds decks in English or French.

## What it does

It opens with a few questions, grouped so you answer them in one go: the output format (Keynote, PowerPoint, Figma Slides, Canva or video), your branding as a design-system link or HEX colors, the type of audience, and the tone. In English that means casual first names or formal titles; in French, *tu* or *vous*. A short brief follows, and you can leave parts of it blank.

Before rendering anything, it shows you the run of show as a table (blocks, speakers, timing and slide count) so you can reorder it while that's still cheap.

You get a 16:9 .pptx with fades and click builds, an HTML preview, and the full script twice: once in the speaker notes and once in `script.md`.

## What it won't do

Make up a number. A missing figure, price, date, customer name or testimonial stays on the slide as `[TO PROVIDE: …]`, dimmed so it reads as an empty slot (`[À FOURNIR]` in a French deck). Wording gets different treatment: when your brief is thin, the skill drafts the thesis or a feature description itself and flags it `[TO VALIDATE: …]` until you sign off.

It also won't use the same tricks on every audience. Charts with no axis labels, figures with no source, an unnamed competitor, "up to" claims: they play fine at a consumer launch and get picked apart by investors or a procurement team. You choose the audience type once, and the skill decides what each slide can get away with.

Type sizes, weights, letter spacing and positions on the slide come from 67 frames pulled from those eight keynotes and measured pixel by pixel. The repo holds no Apple images or fonts, and no quotes from Apple.

## Layouts

The skill knows 25 slide types taken from the keynotes (title card, hero number, bento recap, product alone, screenshot inside a device, full-bleed photo, pricing, availability and others), with 147 layout variants spread across them. Each type has its own entry in `references/designs.md`. The skill picks a layout for every slide and switches things up from one block to the next.

It also hands you a list of the visuals to prepare, each with a brief covering type, framing, background (transparent or not) and resolution. Until a visual shows up, the slide carries a neutral placeholder with no numbers and no logo on it.

![137 of the 147 layouts, rendered by the skill for a fictional brand](skills/sebmakes-keynote-likeapple-pls/assets/galerie.jpg)

## Install

Claude Code, as a plugin:

```bash
claude plugin marketplace add sebmake/sebmakes-keynote-likeapple-pls
claude plugin install sebmakes-keynote-likeapple-pls@sebmake
```

Claude Code, by hand: copy the `skills/sebmakes-keynote-likeapple-pls` folder into `~/.claude/skills/`.

Claude.ai: download the `.skill` file from the latest release on the Releases tab, then upload it in your skill settings.

You'll need Node 18 or later and Python 3.9 or later. The skill runs `npm install` for pptxgenjs on the first render; Python runs the deck check.

## Usage

Ask for it the way you'd ask a colleague:

> Make me a 10-minute Apple-style presentation introducing our new invoicing module to customers. Colors: #091230 and #E4FF3A. Formal tone. PowerPoint.

Two things still need a human eye. Click builds don't always survive an import into Keynote or Canva, and the fonts have to be installed on whatever machine you present from.

---

Not affiliated with, endorsed or sponsored by Apple Inc. Apple and Keynote are trademarks of Apple Inc., registered in the U.S. and other countries. MIT License.
