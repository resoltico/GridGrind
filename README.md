---
afad: "3.5"
version: "0.49.0"
domain: OVERVIEW
updated: "2026-04-20"
route:
  keywords: [gridgrind, excel, xlsx, workbook, automation, spreadsheet, quick-start]
  questions: ["what is gridgrind", "who is gridgrind for", "how do i start with gridgrind", "can gridgrind work with existing excel files", "do i need excel installed"]
---

# GridGrind — a calmer way to keep `.xlsx` workbook work on track

*GridGrind helps you update, check, and read `.xlsx` workbooks in one repeatable run, so spreadsheet work that comes back again and again stops eating manual time.*

## First Sip

- Keep recurring workbook work consistent.
- Update a workbook, check it, and read back what matters in one pass.
- Catch problems before a file is written.
- Start from a quick start guide and real examples instead of a blank page.

## Who It's For

- For teams whose day still runs through Excel workbooks.
- For people tired of opening the same `.xlsx` files and walking through the same steps by hand.
- For workbook work that should be consistent, reviewable, and easy to run again.

## What It Changes

- Instead of opening a workbook and walking through the same steps again, you set the work up once and run it again when it comes back next week, next month, or in the next handoff.
- GridGrind can start from a new workbook or from an existing `.xlsx` file.
- When a run passes, you keep the workbook and the facts you asked for. When it fails, you get a clear error instead of a partly written file.

## Proof and Trust

- GridGrind is built on Apache POI XSSF, a long-established Java `.xlsx` library, instead of a home-grown spreadsheet file layer.
- If a run fails, GridGrind stops before writing the workbook, so you do not get a partly written file.
- GridGrind ships a user-facing quick start, real example files, and supporting docs when you want more detail.
- The limits are explicit instead of hand-wavy: low-memory `STREAMING_WRITE` is intentionally narrow and uses `ENSURE_SHEET`, `APPEND_ROW`, and optional `execution.calculation.markRecalculateOnOpen=true`.
- Formula boundaries are explicit too: request-authored array braces, `LAMBDA`, and `LET` are rejected as `INVALID_FORMULA` today because Apache POI cannot parse them yet.
- Limits and format boundaries are documented up front in [docs/LIMITATIONS.md](docs/LIMITATIONS.md).
- The scope is explicit: GridGrind is for `.xlsx` workbooks, while `.xls`, `.xlsm`, and `.xlsb` are out of scope.

## Start Here

- New to GridGrind: [docs/QUICK_START.md](docs/QUICK_START.md)
- Want a concrete example first: [examples/budget-request.json](examples/budget-request.json)
- Want a no-save health-check example: [examples/workbook-health-request.json](examples/workbook-health-request.json)
- Want the reference docs: [docs/QUICK_REFERENCE.md](docs/QUICK_REFERENCE.md), [docs/OPERATIONS.md](docs/OPERATIONS.md), and [docs/ERRORS.md](docs/ERRORS.md)
- Want the runnable download: [latest release](https://github.com/resoltico/GridGrind/releases/latest)

## Questions Before You Pour

### Do I need Excel installed?

No. GridGrind reads and writes `.xlsx` workbooks itself. You can run it without Microsoft Excel on the machine.

### Can I use existing workbooks?

Yes. GridGrind can start from a new workbook or open an existing `.xlsx` file.

### Do I need to start from scratch?

No. GridGrind ships example files and a quick start, so you can start from something close to your own workbook job and adjust it.

### What kinds of workbook work fit best?

GridGrind is strongest when the same `.xlsx` work should run the same way each time: filling workbook content, checking results, reading facts back, and running review or health-check passes.
For the detailed supported surface, see [docs/OPERATIONS.md](docs/OPERATIONS.md).

### What happens if something fails?

GridGrind returns a clear error response and skips workbook persistence.
It does not save a partial `.xlsx` file after a failed run.

### Is it right for every kind of spreadsheet work?

No. GridGrind is best for repeat `.xlsx` work that should run the same way each time.
If your work depends on `.xls`, `.xlsm`, `.xlsb`, or one-off spreadsheet work, it is not the right fit.

---

## Legal

GridGrind is released under the MIT License. Bundled third-party notices live in [NOTICE](NOTICE), and patent considerations are documented in [PATENTS.md](PATENTS.md).

[LICENSE](LICENSE) | [NOTICE](NOTICE) | [PATENTS.md](PATENTS.md)
