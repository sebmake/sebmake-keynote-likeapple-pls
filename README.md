
<a href="skills/sebmake-keynote-likeapple-pls/assets/galerie.jpg">
<img width="360" alt="137 of the 148 slide layouts, rendered by the skill for a made-up brand" src="skills/sebmake-keynote-likeapple-pls/assets/galerie.jpg" />
</a>

# sebmake-keynote-likeapple-pls

For anyone who has to present: a 16:9 deck built the way Apple builds its keynotes, plus the script you read over it.

Making a deck that feels like an Apple keynote is hard. Big white type on black isn't enough, because what makes those keynotes work is the order of the news, the pacing, and one idea per slide.

This skill is based on eight public Apple keynotes, from 2019 to 2026. They were studied slide by slide, and 67 frames were measured pixel by pixel for type size, weight, letter spacing and position.

Your deck keeps your colors, your fonts, your product and your numbers. This repo holds no Apple images, fonts or quotes.

The skill's instructions are in French, but it builds decks in English or French.

[Get the latest release](https://github.com/sebmake/sebmake-keynote-likeapple-pls/releases/latest)

## Install

```bash
npx skills@latest add sebmake/sebmake-keynote-likeapple-pls
```

It's one of the four [LikeApplePLS](https://github.com/sebmake/skills) skills, and `npx skills@latest add sebmake/skills/LikeApplePLS` installs all of them.

Or as a Claude Code plugin:

```bash
claude plugin marketplace add sebmake/sebmake-keynote-likeapple-pls
claude plugin install sebmake-keynote-likeapple-pls@sebmake
```

You can also copy `skills/sebmake-keynote-likeapple-pls` into `~/.claude/skills/` by hand. On Claude.ai, upload the `.skill` file from the latest release in your skill settings.

You'll need Node 18 or later and Python 3.9 or later. The first render runs `npm install` for pptxgenjs.

## Why use it?

Agents copy the look of a keynote, not the way it's built.

Ask an agent for an "Apple-style" deck and you'll get huge titles on black. In a side-by-side test, the agent working without the skill got that part right, then made up amounts to fill the empty slides.

This skill saves the biggest news for last and walks each announcement through set steps. It picks from 25 slide types and 148 layout variants, and it changes the layout from one block to the next. A missing figure, price, date or customer name stays on the slide as `[TO PROVIDE: …]`, dimmed so it reads as an empty slot. When your brief is thin, it drafts the wording itself and flags it `[TO VALIDATE: …]` until you sign off.

Some tricks only work on some audiences. A chart with no axis labels, an unsourced figure or an "up to" claim works at a consumer launch, but investors or a buying team will question it. You pick the audience once, and the skill decides what each slide can show.

Before anything renders, you see the run of show (blocks, speakers, timing, slide count) and can reorder it before anything is built. Then you get a 16:9 .pptx with fades and click builds, an HTML preview, and the full script, both in the speaker notes and in `script.md`.

## Reference

- **[SKILL.md](./skills/sebmake-keynote-likeapple-pls/SKILL.md)** — The main skill. It asks for the format (Keynote, PowerPoint, Figma Slides, Canva or video), your colors, the audience and the tone, then writes the run of show, the slides and the script.
- **[grammaire.md](./skills/sebmake-keynote-likeapple-pls/references/grammaire.md)** — How the keynotes are put together: run of show, announcement steps, hand-offs between speakers, recaps.
- **[designs.md](./skills/sebmake-keynote-likeapple-pls/references/designs.md)** — One entry per slide type, with its layout zones, its variants and the visual to ask you for.
- **[style-visuel.md](./skills/sebmake-keynote-likeapple-pls/references/style-visuel.md)** — The measured type sizes, weights, spacing and positions.
- **[publics.md](./skills/sebmake-keynote-likeapple-pls/references/publics.md)** — What each type of audience will accept on a slide.
- **[deck-spec.md](./skills/sebmake-keynote-likeapple-pls/references/deck-spec.md)** — The `deck.json` format, with a full example.

The reference files are in French.

Two things still need checking by hand: click builds don't always survive an import into Keynote or Canva, and your fonts have to be installed on the machine you present from.

---

Not affiliated with, endorsed or sponsored by Apple Inc. Apple and Keynote are trademarks of Apple Inc., registered in the U.S. and other countries. MIT License.
