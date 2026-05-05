# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Nature

This is a **documentation / research repository**, not a software project. It accompanies the paper *"Dive into Claude Code: The Design Space of Today's AI Agent System"* (an architectural analysis of Claude Code v2.1.88). There is **no source code, no build system, no tests, and no package manager**. All work in this repo is editing Markdown, managing the PDF paper, and curating diagrams/assets.

Do not invent build/lint/test commands. If asked to "run the project," clarify with the user — there is nothing to run.

## Layout and Conventions

```
README.md           README_zh.md          ← top-level landing page (EN + ZH mirror)
docs/architecture.md           docs/architecture_zh.md
docs/build-your-own-agent.md   docs/build-your-own-agent_zh.md
docs/related-resources.md      docs/related-resources_zh.md
assets/*.png        ← diagrams referenced from the Markdown
paper/Dive_into_Claude_Code.pdf
CITATION.cff
.github/workflows/refresh-star-history.yml   ← the only automation; hourly camo purge
```

**Bilingual mirroring is a hard invariant.** Every English document has a `_zh.md` Chinese sibling that must be kept in sync. When editing `README.md`, edit `README_zh.md` in the same change; same for every file under `docs/`. Section anchors, table rows, asset paths, and image references must match across both languages — the language toggle at the top of each file (`English | 中文`) and "Back to top" anchors depend on it.

**Asset references resolve relative to the file.** `README.md` uses `./assets/foo.png`; files under `docs/` use `../assets/foo.png`. When adding a diagram, place the PNG in `assets/` and reference it from both the EN and ZH versions.

**README uses collapsible `<details>` sections** (one is `<details open>` — "Architecture at a Glance"). Preserve this pattern when adding new top-level sections; don't convert them to plain headings.

**Numbers in the prose are load-bearing claims** tied to the paper (e.g. "1.6% AI / 98.4% infrastructure", "7 safety layers", "5 compaction stages", "54 tools", "27 hook events", "v2.1.88", "~512K lines", "1,884 files"). Don't paraphrase them away or change a count without checking the paper PDF and the corresponding number in the Chinese README.

## Things That Look Wrong But Aren't (Yet)

- `CITATION.cff` contains literal `"To be updated"` placeholders for authors and a `# TODO: Update authors` comment. The BibTeX block in `README.md` already lists "Jiacheng Liu, Xiaohan Zhao, Xinyi Shang, and Zhiqiang Shen". If asked to fill the citation, propagate those names to `CITATION.cff` (both `authors:` blocks) and verify against `README_zh.md`.
- The arXiv ID `2604.14228` and the year `2026` are intentional (the paper is dated 2026/04/11 per `CITATION.cff`); don't "correct" them to a current year.
- The README links to many third-party repos with star badges. Most are real — but if the user asks you to add a new project to a comparison table or list, fetch and read the target repo before writing source-grounded prose; that's the editorial standard the existing tables follow ("Cells are source-grounded; this is not a feature scoreboard").

## License

Content is **CC BY-NC-SA 4.0**. Material copied/adapted from this repo elsewhere must carry the same license; any new prose added here inherits it.
