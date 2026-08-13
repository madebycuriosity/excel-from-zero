# Excel From Zero

**An interactive, beginner-first Excel learning platform for people who have never opened a spreadsheet.**

> **Nobody ever showed you the little square.**

Excel is usually taught starting from formulas and functions. For a complete beginner the real obstacle is much earlier than that: what a cell is, where to click, what happens when you type, how a spreadsheet actually behaves.

**Excel From Zero** makes those first steps practical, interactive and less intimidating.

## Live Demo

**[Open Excel From Zero](https://nobodytoldmeanything.netlify.app/)**

## What Is It?

A hands-on learning environment where people practise in a real spreadsheet grid instead of watching videos.

- **13 parts, 52 lessons**
- A working spreadsheet engine, not a mock-up
- A diagnostic that places you at the right lesson
- Feedback that responds to *how* an answer is wrong
- Progress saved between visits
- No account, no download, no cost

## Learning Pathway

**Part 1 — Getting Comfortable**
Cells, addresses, typing, editing, undoing, selecting ranges, and what `#####` means.

**Part 2 — Making It Calculate**
Formulas, `SUM`, relative and absolute references, and the fill handle.

**Part 3 — What Kind of Data Is This?**
Categorical and numerical, nominal and ordinal, interval and ratio.

**Part 4 — Describing Data**
Mean, median, outliers, range, standard deviation, quartiles, and choosing the right measure.

**Part 5 — Working With Real Data**
Sorting, `COUNTIF`, `SUMIF`, `XLOOKUP`, PivotTables.

**Part 6 — Showing What You Found**
Choosing an appropriate chart, and spotting a misleading axis.

**Part 7 — Things That Go Wrong Quietly**
Numbers stored as text, ambiguous dates, disappearing zeros, formatting, Paste Values, incomplete ranges.

**Part 8 — Making Excel Decide**
`IF`, `AND`, `OR`, nested `IF`.

**Part 9 — Put It Together**
Two projects with no real instructions: diagnose a broken spreadsheet, then say something useful about it.

**Part 10 — Text, Dates and Errors**
`TRIM`, `LEN`, `LEFT`, `MID`, `RIGHT`, `DAY`, `MONTH`, `YEAR`, `IFERROR`.

**Part 11 — Numbers That Mean Something**
Percentage change, `ROUND`, `CORREL`.

**Part 12 — The Menu, Not the Formula**
Filter, Remove Duplicates, Find and Replace, Conditional Formatting, Data Validation, Text to Columns.

**Part 13 — The Judgement Calls**
Missing values, and when the mode is the only honest average.

## Design Philosophy

**Learn by doing, not by watching.**

Every lesson ends with a definite capability — something the learner can now do that they could not do before.

Two principles shape the whole thing. Assume nothing: the first lessons cover clicking, typing and the fill handle, because nobody is ever shown these and everyone is assumed to know them. And teach from mistakes: when an answer is wrong in a specific, common way, the response explains what that particular mistake reveals rather than saying "incorrect".

## Who Is It For?

Australian undergraduate students starting from zero with Excel, and anyone else who needs a practical introduction to spreadsheet work — career changers, people returning to work, small business owners doing their own books.

## How It Works

A single `index.html` file. No dependencies, no build step, no backend. Open it locally or drop the folder on any static host.

**A spreadsheet engine.** Cells hold values or formulas, formulas evaluate against other cells, results update when inputs change. Around 40 functions:

`SUM` `AVERAGE` `COUNT` `COUNTA` `MIN` `MAX` `MEDIAN` `MODE.SNGL` `STDEV.S` `STDEV.P` `VAR.S` `QUARTILE.INC` `PERCENTILE.INC` `LARGE` `SMALL` `CORREL` `ROUND` `ABS` `SQRT` `IF` `AND` `OR` `NOT` `IFERROR` `COUNTIF` `SUMIF` `AVERAGEIF` `XLOOKUP` `VLOOKUP` `TRIM` `LEN` `LEFT` `RIGHT` `MID` `UPPER` `LOWER` `PROPER` `DAY` `MONTH` `YEAR`

Plus operator precedence, brackets, exponents, percent literals, `&` concatenation, comparison operators, reference shifting including mixed forms like `B$2`, and Excel's error values (`#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?`, `#REF!`).

Behaviour is matched to Excel rather than approximated: text left-aligned and numbers right-aligned, text that looks numeric excluded from `SUM`, case-insensitive text comparison, `AVERAGE` ignoring text and blanks but counting zero, and copying a formula shifting its references while cutting one does not.

**Direct manipulation.** Click and drag to select ranges, type straight into cells, F2 or double-click to edit in place, drag the fill handle, drag column borders to resize, Ctrl+Z / C / X / V, and Excel-style point mode — type `=SUM(`, drag across the cells, and the reference writes itself into the formula.

**Tools, where a lesson needs one.** Sort, Filter, Remove Duplicates, Find and Replace, Conditional Formatting, Data Validation, Text to Columns, number formatting, clipboard, a PivotTable builder, and an SVG chart renderer (column, line, pie, scatter, histogram).

## What This Is Not

A spreadsheet emulator. Deliberately.

Not implemented: multiple sheets, the ribbon, right-click menus, inserting or deleting rows and columns, AutoSum, saving files, whole-column references like `B:B`, date arithmetic. The grid is six columns wide.

The reasoning: **the thinking transfers, the menus don't.** References, ranges, what a function is, why an average can mislead, when a chart lies — these are identical in every spreadsheet program and every version. Where Remove Duplicates sits in the ribbon changes between releases and differs on Mac.

Someone finishing all 52 lessons will open real Excel knowing what they are trying to do, and will still spend a minute finding the button.

## Adding a Lesson

Lessons live in the `LESSONS` array:

```js
{
  title: 'SUM a range',
  sub:   'for when there are more than two',
  task:  'In B8, total all six weeks using SUM.',
  teach: 'Explanation. <b>Bold</b> is allowed.',
  hint:  'Shown when the learner presses "Stuck?"',
  tool:  'filter',        // optional: sort filter dedupe findreplace
                          // condformat validation split format clip pivot chart
  setup: function () {
    grid(9, 5, { A1: 'Week', B1: 'Sales', B2: '1240' });
    S.widths = { A: 110 };
    sel.a = sel.b = 'B8';
  },
  check: function () {
    if (raw('B8') === '') return [0, ''];
    if (val('B8') === 7195) return [1, 'Success message'];
    return [0, 'Nudge toward the answer'];
  }
}
```

`check()` returns `[code, message]`:

| code | meaning | styling |
|------|---------|---------|
| `1` | correct — unlocks Next, marks the lesson complete | green |
| `0` | not yet — a nudge, or `''` to stay silent | amber |
| `-1` | a *specific* wrong answer worth teaching from | red |

**`-1` is where the value is.** Detect the particular mistake and explain what it reveals. Existing examples: `=340+210` instead of `=B2+B3`, `>` instead of `>=`, `OR` instead of `AND`, grade bands tested in the wrong order, sorting one column instead of the table.

Helpers available inside `setup` and `check`:

| helper | does |
|--------|------|
| `grid(rows, cols, cells)` | builds the sheet; values are strings, or `{v:'420', f:'aud'}` to format |
| `raw(k)` | literal cell contents, formula included |
| `val(k)` | evaluated result |
| `numOf(k)` | numeric value, or `null` if it isn't a number |
| `sel.a`, `sel.b` | current selection anchor and focus |
| `T` | tool state — `T.chartType`, `T.pivRow`, `T.filterOn`, `T.sortMode` |
| `S.cells[k].f` | number format: `'aud'`, `'pct'`, `'num'` |

Prefix a value with an apostrophe (`"'0800"`, `"'120"`) to force it to text, exactly as Excel does. That is how the numbers-stored-as-text and leading-zero lessons work.

After adding a lesson, update `MODULES` — each entry has `from` and `to` indexes into `LESSONS`, so inserting mid-array shifts everything after it.

## Future Development

- Mobile: fill-handle dragging is fiddly on touch
- Accessibility: keyboard-only grid navigation and screen reader support are incomplete
- Real date handling, so dates can be subtracted
- Splitting the single file into modules as it grows
- More project-based activities

## Project Status

**Prototype, in active development.** An ongoing experiment in interactive learning, user experience and product design.

---

**Excel From Zero — a practice grid, not a video course.**
