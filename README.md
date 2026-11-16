# Folder Triage

**Open a folder full of files, look at them one by one, press one key, and they file themselves away.**

Folder Triage is a tiny, dependency-free Windows tool for the boring part of any review job: opening dozens (or hundreds) of documents one at a time, deciding where each one belongs, closing it, and dragging it into the right folder. This tool does the opening, closing and moving for you — you only look and press `1`, `2` or `3`.

It is **human-in-the-loop by design**: no AI, no cloud, no uploading your files anywhere. It just drives the applications you already have installed.

```
┌──────────────── 你在这里看/敲键 ────────────────┐   ┌──────── 自动打开的作品 ────────┐
│ [3/87] 校园PPT模板设计大赛+学术科创组+222      │   │                                │
│        文件：作品.pptx   12 页 / 4.3 MB        │   │         (in your default app)  │
│        本次：好 1  中 1  坏 0                   │   │                                │
│                                                │   │                                │
│        分类 _                                   │   │                                │
└────────────────────────────────────────────────┘   └────────────────────────────────┘
```

---

## What it is good for

Any task that is "same action, many items":

| Scenario | Categories |
|---|---|
| 参赛作品 / 投稿初评 | 好 / 中 / 坏 |
| 照片分选、素材归档 | 留 / 修 / 删 |
| 报销单据、合同整理 | 已核 / 待补 / 作废 |
| 简历初筛 | 面试 / 待定 / 不合适 |
| 收集的 PDF、PPT 分类 | any N buckets |

The category names are configurable, and the count of categories is arbitrary (`1..9`).

## Features

- **Works with any file type and any viewer** — it just launches the file with its default application and tracks the window that appears. PowerPoint, WPS, LibreOffice, Acrobat, Photos, a video player: whatever Windows uses.
- **One keystroke per item** — `1/2/3…` assigns a category, then the item is moved and the next one opens.
- **Preview mode** — optionally opens a random 1/3 of the items first, so you build a consistent standard before you start judging.
- **Resume anywhere** — progress lives in `_triage_state.json`; quit at any time and rerun.
- **Undo** — `triage.cmd -Undo` moves everything back, in reverse order.
- **Safe moves** — atomic directory/file renames, name-collision handling, retries while a viewer still holds the file, and no silent failures.
- **CSV log** — `_triage_log.csv` (time, name, verdict, note), opens straight in Excel.
- **No install** — a single PowerShell script + a `.cmd` launcher. Uses the Windows built-in PowerShell 5.1 if PowerShell 7 is absent.

## Quick start

1. Download and unzip the [latest release](../../releases/latest).
2. Drop `triage.ps1` and `triage.cmd` into the folder that **directly contains** the items:

```
D:\review\
  ├─ team-111\          ← each folder is one item (its largest matching file is opened)
  ├─ team-222\
  ├─ loose-deck.pptx    ← loose files work too
  ├─ triage.cmd
  └─ triage.ps1
```

3. Double-click `triage.cmd`.

It asks once whether to preview a random 1/3 first, then starts.

## Keys

| Key | Action |
|---|---|
| `1` `2` `3` … | put the item into category N |
| `Enter` | put the item into category 1 |
| `s` | skip (leave in place, just log it) |
| `r` | reopen the current item |
| `o` | reveal the current item in Explorer |
| `k` | reset: force-close the viewer and reopen the item |
| `q` | quit (progress is saved) |

## Options

```powershell
triage.cmd                                  # normal run
triage.cmd -NoPreview                       # skip the preview question
triage.cmd -Start 12                        # resume at item 12
triage.cmd -Categories 留,修,删              # custom category names
triage.cmd -Any                             # accept every file type
triage.cmd -Exts pdf,docx,pptx              # restrict file types
triage.cmd -Undo                            # undo the last run's moves
triage.cmd -Path D:\other\folder            # run against another folder
```

## Generated files

| File | Meaning |
|---|---|
| `好\` `中\` `坏\` | category folders, created on first use |
| `_triage_state.json` | progress; delete it to start over |
| `_triage_log.csv` | per-item log |

Folders and files starting with `_` (and the category folders themselves) are ignored when scanning.

## Notes and limitations

- Windows only. The layout helper is optional (`-PinConsole`); by default the script never touches the terminal window it runs in.
- Items are considered to be the **direct children** of the working folder: a subfolder is one item, and the largest matching file inside is what gets opened.
- Files that are still held open by a viewer are retried; if the move still fails it is reported instead of silently skipping.
- The console is only pinned/shrunk when started through `triage.cmd` and only the console that belongs to this process — never another window.

## 中文速览

把一个装满作品的文件夹交给它，它逐个用默认程序打开，你看完敲 `1/2/3`，它自动关掉、把该文件夹（或文件）挪进对应分类，再打开下一个。不联网、不装环境、不改原文件内容。

- 分类名可改、数量可改（`-Categories 好,中,坏`）
- 启动可先随机预览 1/3，帮助建立统一标准
- 断点续评、可撤销（`-Undo`）、有 CSV 明细
- 支持任意文件类型：PPT/PDF/Word/图片/视频都行

## License

MIT
