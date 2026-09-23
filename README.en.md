# Sublyric · Excel Subtitle Toolkit

English | [简体中文](README.md)

> Clean up, proofread, and export subtitles — all inside Excel.
> Free and open-source. Once you're done, drop the exported SRT straight
> into [Sublyric Subtitle](https://shinesoft.cn/zh-CN/) for timeline
> fine-tuning and karaoke effects.

![Project banner](请在此插入项目封面图/Banner)

---

## Table of Contents

- [1. Features](#1-features)
  - [1.1 Why this project exists](#11-why-this-project-exists)
  - [1.2 What problems it solves](#12-what-problems-it-solves)
  - [1.3 Feature overview](#13-feature-overview)
- [2. Installation](#2-installation)
- [3. About Sublyric Subtitle / the team](#3-about-sublyric-subtitle--the-team)

---

## 1. Features

### 1.1 Why this project exists

This toolkit started out as an internal script our team used alongside
[Sublyric Subtitle](https://shinesoft.cn/zh-CN/) — a small Excel VBA macro
for handling subtitles. Over time we noticed that subtitle groups and
localization teams spend most of their prep time in the most basic office
software: pasting text, aligning timecodes, catching typos, reading things
back to check they sound right. Rather than have everyone maintain their
own slightly-different copy of the macro, we turned it into a single
installable add-in and open-sourced it — use it as-is, or fork it and make
it your own.

### 1.2 What problems it solves

A few recurring headaches during subtitle/script prep:

- **Messy bilingual formatting**: whether to add spaces between Chinese and
  English text, mixed full-width/half-width punctuation, trailing
  punctuation that should or shouldn't be there — fixing this line by line
  by hand takes forever
- **Proofreading by eye only**: once the text is edited, catching awkward
  phrasing, missing words, or bad line breaks is hard to do visually;
  having someone read it aloud (or doing it yourself) works but breaks your
  flow
- **Tedious format conversion**: turning a bilingual table in Excel into a
  proper SRT file, or a clean Word document for review, is easy to get
  wrong when done by hand

### 1.3 Feature overview

Adds a "Sublyric Subtitle Tools" tab to the Excel ribbon:

![Excel ribbon screenshot](请在此插入Excel功能区截图)

| Group | Feature | Description |
|---|---|---|
| Formatting | Basic formatting | One click to style the header row, add alternating row shading, normalize the time columns, and clean up spacing/casing |
| Formatting | Remove trailing punctuation | Strips extra punctuation at the end of a line (question marks and exclamation points are kept) |
| Formatting | Add CJK/English spacing | Automatically inserts or removes spaces between Chinese, English, numbers, and punctuation |
| Formatting | Read-aloud proofreading | Select text and it's read aloud using Edge's natural voice — starts automatically, no second click needed |
| Export | Export bilingual subtitles | Exports an SRT using the column layout: index / start time / end time / English / Chinese |
| Export | Export Chinese-only subtitles | Exports just the Chinese lines as an SRT |
| Export | Export bilingual document | Exports a Word document for easy review and sharing |

**Read-aloud proofreading preview**

![Read-aloud window screenshot](请在此插入朗读窗口截图/GIF)

Read-aloud proofreading calls the real Microsoft Edge browser installed on
your system (Chromium-based), using Edge's natural voices rather than
old-school robotic TTS. The window docks to the right edge of the screen
and fills the full height, giving something close to a sidebar feel. Select
text and it starts reading immediately; select a new piece of text and it
swaps in and keeps reading — no need to close and reopen anything.

---

## 2. Installation

### Option A: One-click install via PowerShell (recommended)

Follow [`docs/01-Excel打包指南.md`](docs/01-Excel打包指南.md) to build
`ExcelToSRT.xlam` from source, then install it along with `reader.html`
using the scripts under `install/`:

```powershell
# Right-click install-excel.ps1 → "Run with PowerShell"
# or, from the command line:
powershell -ExecutionPolicy Bypass -File .\install-excel.ps1
```

After installation, reopen Excel — the "Sublyric Subtitle Tools" tab will
appear automatically, and it will keep loading every time you open Excel
from then on, with no further setup needed.

To uninstall, run `uninstall-excel.ps1`. Everything happens at the current
Windows user level, no admin rights required, and uninstalling leaves no
leftovers.

See [`docs/03-PowerShell使用指南.md`](docs/03-PowerShell使用指南.md) for
details (currently Chinese-only — translation welcome via PR).

![Post-install screenshot](请在此插入安装完成后的功能区截图)

### Option B: Build from source yourself

The source lives under `excel-addin/`. Follow
[`docs/01-Excel打包指南.md`](docs/01-Excel打包指南.md) to build it from
scratch — it also covers how to wire up custom icons and the Ribbon XML.

---

## 3. About Sublyric Subtitle / the team

This toolkit is built and maintained by the **Sublyric Subtitle** team.

**Sublyric Subtitle** is an AI-powered subtitling app covering the full
workflow: **speech-to-text → translation → fine-tuning → karaoke effects →
export**.

- AI transcription with automatic timeline alignment, plus online
  translation into 19+ languages
- Multi-track subtitles synced with an audio waveform, CPS checks, and
  millisecond-level timing adjustments
- 50+ animated karaoke lyric presets, with AI-assisted syllable splitting
  (K-Cut)
- Exports to MP4/MKV/MOV video, and SRT/ASS/WebVTT/Word/Excel and more
  (10+ formats)

![Sublyric Subtitle product screenshot](请在此插入韵词字幕软件截图)

- 🌐 Website: https://shinesoft.cn/zh-CN/
- 📥 Free download: https://shinesoft.cn/zh-CN/download/
- 📖 Tutorials & case studies: https://shinesoft.cn/zh-CN/case-study/

This Excel toolkit is meant as a "prep-stage helper" — it smooths out the
text-editing work you'd do in Excel before the real subtitling begins.
The SRT it exports can be dropped straight into Sublyric Subtitle for
timeline fine-tuning and effects work.

---

## 🙌 Contributing

Found a bug or have an idea? Open an [issue](../../issues), or fork the
repo and send a pull request. If this tool has been useful to you, check
out [Sublyric Subtitle](https://shinesoft.cn/zh-CN/) itself, and consider
leaving a ⭐.

## License

MIT — free to use, modify, and redistribute.
