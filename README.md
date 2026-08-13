# README additions and corrections

Sections below are ready to paste into `README.md`. Three replace existing text; three are new.

---

## REPLACE the "Current Technology" section with this

## How It Works

A single `index.html` file with no dependencies, no build step and no backend. Open it locally or drop the folder on any static host.

Inside it are three things worth knowing about.

**A spreadsheet engine.** Not a mock-up of one. Cells hold values or formulas, formulas evaluate against other cells, and the results update when their inputs change. Around 40 functions are implemented:

`SUM` `AVERAGE` `COUNT` `COUNTA` `MIN` `MAX` `MEDIAN` `MODE.SNGL` `STDEV.S` `STDEV.P` `VAR.S` `QUARTILE.INC` `PERCENTILE.INC` `LARGE` `SMALL` `CORREL` `ROUND` `ABS` `SQRT` `IF` `AND` `OR` `NOT` `IFERROR` `COUNTIF` `SUMIF` `AVERAGEIF` `XLOOKUP` `VLOOKUP` `TRIM` `LEN` `LEFT` `RIGHT` `MID` `UPPER` `LOWER` `PROPER` `DAY` `MONTH` `YEAR`

Alongside operator precedence, brackets, exponents, percent literals, `&` concatenation, comparison operators, relative and absolute reference shifting including mixed forms like `B$2`, and Excel's error values (`#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?`, `#REF!`).

Behaviour is matched to Excel rather than approximated. Text is left-aligned and numbers right-aligned; text that looks numeric is excluded from `SUM`; text comparison is case-insensitive; `AVERAGE` ignores text and blanks but counts zero; copying a formula shifts its references while cutting one does not.

**Direct manipulation.** Click and drag to select ranges, type straight into cells, F2 or double-click to edit in place, drag the fill handle, drag column borders to resize, Ctrl+Z, Ctrl+C, Ctrl+X, Ctrl+V, and Excel-style point mode — type `=SUM(`, drag across the cells, and the reference writes itself into the formula.

**Tools, where a lesson needs one.** Sort, Filter, Remove Duplicates, Find and Replace, Conditional Formatting, Data Validation, Text to Columns, number formatting, clipboard, PivotTable builder, and an SVG chart renderer (column, line, pie, scatter, histogram).

---

## REPLACE the "Future Development" section with this

## What's Already Built

- Progress saved across sessions, with a clear-progress link
- A diagnostic that places a learner at the right lesson
- Feedback that responds to *how* an answer is wrong, not just that it is
- Module structure with per-part progress and completion summaries

## Future Development

- Mobile: fill-handle dragging is fiddly on touch
- Accessibility: keyboard-only grid navigation and screen reader support are incomplete
- Splitting the single file into modules as it grows
- More project-based activities
- Real date handling (dates are currently read as text, so date arithmetic isn't possible)

---

## ADD this section — sets expectations honestly

## What This Is Not

A spreadsheet emulator. Deliberately.

Not implemented: multiple sheets, the ribbon, right-click menus, inserting or deleting rows and columns, AutoSum, saving files, whole-column references like `B:B`, and date arithmetic. The grid is six columns wide.

The reasoning: **the thinking transfers, the menus don't.** References, ranges, what a function is, why an average can mislead, when a chart lies — these are the same in every spreadsheet program and every version. Where Remove Duplicates sits in the ribbon changes between releases and differs on Mac.

A learner finishing all 52 lessons will open real Excel knowing what they are trying to do, and will still spend a minute finding the button.

---

## ADD this section — so you can add lessons in six months without re-reading the file

## Adding a Lesson

Lessons live in the `LESSONS` array. Each is an object:

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
    if (raw('B8') === '') return [0, ''];                 // nothing yet, stay quiet
    if (val('B8') === 7195) return [1, 'Success message'];
    return [0, 'Nudge toward the answer'];
  }
}
```

`check()` returns `[code, message]`:

| code | meaning | styling |
|------|---------|---------|
| `1` | correct — unlocks Next, marks complete | green |
| `0` | not yet — a nudge, or `''` to show nothing | amber |
| `-1` | a *specific* wrong answer worth teaching from | red |

**`-1` is where the value is.** Detect the particular mistake and explain what it reveals, rather than saying "incorrect". Examples in the file: `=340+210` instead of `=B2+B3`; `>` instead of `>=`; `OR` instead of `AND`; grade bands tested in the wrong order; sorting a single column instead of the table.

Helpers available inside `setup` and `check`:

| helper | does |
|--------|------|
| `grid(rows, cols, cells)` | builds the sheet; values are strings, or `{v:'420', f:'aud'}` for formatting |
| `raw(k)` | the literal contents of a cell, formula included |
| `val(k)` | the evaluated result |
| `numOf(k)` | numeric value, or `null` if it isn't a number |
| `sel.a`, `sel.b` | the current selection anchor and focus |
| `T` | tool state — `T.chartType`, `T.pivRow`, `T.filterOn`, `T.sortMode` and so on |
| `S.cells[k].f` | number format: `'aud'`, `'pct'`, `'num'` |

Prefix a value with an apostrophe (`"'0800"`, `"'120"`) to force it to be text, exactly as Excel does. That's how the numbers-stored-as-text and leading-zero lessons work.

After adding a lesson, update the `MODULES` array — each entry has `from` and `to` indexes into `LESSONS`, so inserting mid-array shifts everything after it.

---

## ADD a LICENSE decision

A public repo with no `LICENSE` file means others may read the code but have no legal right to reuse it. That's the restrictive default and is probably what you want if there's any prospect of commercialising this.

If you'd rather people freely adapt it, MIT is the conventional choice. Either way it should be a decision rather than an oversight.

