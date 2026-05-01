---
afad: "4.0"
version: "0.62.0"
domain: JAVA_AUTHORING
updated: "2026-05-01"
route:
  keywords: [gridgrind, java, authoring, gridgrindplan, targets, values, tables, links, workbookplan, executioninputbindings]
  questions: ["can i use gridgrind from java", "how do i author gridgrind workflows in java", "what is gridgrindplan", "how do i execute a java-authored plan in process", "how do source-backed inputs work from java"]
---

# Java Authoring Guide

**Purpose**: Explain the fluent Java authoring layer that emits the same canonical `WorkbookPlan`
as the JSON protocol and intentionally stops at that contract boundary.
**Example**: [../examples/java-authoring-workflow.java](../examples/java-authoring-workflow.java)
**Underlying contract**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md)
and [OPERATIONS.md](./OPERATIONS.md)

GridGrind has two first-class authoring surfaces:

- JSON request documents for the CLI, fat JAR, and Docker image
- the `dev.erst.gridgrind.authoring` module for selector-first fluent Java workflows

Both surfaces compile to the same canonical `WorkbookPlan`. The Java layer is not a separate
execution engine and does not bypass request validation, execution-mode rules, or workbook safety
checks. It authors the canonical plan; execution stays explicit through the executor layer when
you choose to run in process.

Within the repository's JPMS layout, `dev.erst.gridgrind.authoring` requires transitive
`dev.erst.gridgrind.contract`. The executor remains a separate dependency so applications that
only want to emit plan JSON do not inherit the in-process runtime graph.

## Authoring Types

| Type | Owns |
|:-----|:-----|
| `GridGrindPlan` | Start a plan with `newWorkbook()`, `open(path)`, or `from(plan)`; set persistence and journal detail; emit the canonical `WorkbookPlan` or JSON |
| `Targets` | Selector builders and fluent target-scoped mutation, inspection, and assertion builders |
| `Values` | Typed cell values, expected values, comments, and source-backed UTF-8 helpers |
| `Tables` | Focused table definitions and style wrappers |
| `Links` | URL, email, file, and document hyperlink targets |

`Checks` and `Queries` still exist internally, but they are not part of the public authoring API.
The public Java surface stays selector-first and hides raw request/response DTO plumbing such as
`TableInput`, `HyperlinkTarget`, and the report namespaces under
`GridGrindWorkbookSurfaceReports`, `GridGrindLayoutSurfaceReports`,
`GridGrindSchemaAndFormulaReports`, and `GridGrindAnalysisReports`.

## Optional Explicit Execution Types

If your application also chooses to execute authored plans in process, the explicit executor layer
adds these types:

| Type | Owns |
|:-----|:-----|
| `ExecutionInputBindings` | Working directory plus optional `STANDARD_INPUT` bytes for explicit in-process execution |
| `GridGrindRequestExecutor` | Transport-neutral execution port; `DefaultGridGrindRequestExecutor` is the production implementation |

Unnamed authored steps receive generated ids such as `mutation-001`, `inspection-001`, and
`assertion-001`. If you need stable caller-owned ids, call `.named("...")` on a
`PlannedMutation`, `PlannedInspection`, or `PlannedAssertion` before adding it to the plan.

## Compile-And-Execution-Verified Example

The checked-in example under
[../examples/java-authoring-workflow.java](../examples/java-authoring-workflow.java) is compiled
and executed by `:authoring-java:test`. A shortened excerpt:

```java
GridGrindPlan plan =
    GridGrindPlan.newWorkbook()
        .saveAs(workspace.resolve("budget.xlsx"))
        .journal(ExecutionJournalLevel.VERBOSE)
        .mutate(Targets.sheet("Budget").ensureExists())
        .mutate(
            Targets.range("Budget", "A1:B3")
                .setRows(
                    List.of(
                        Values.row(Values.text("Item"), Values.text("Amount")),
                        Values.row(Values.text("Hosting"), Values.number(100.0)),
                        Values.row(Values.text("Travel"), Values.number(50.0)))))
        .mutate(
            Targets.tableOnSheet("BudgetTable", "Budget")
                .define(
                    Tables.define(
                        "BudgetTable",
                        "Budget",
                        "A1:B3",
                        false,
                        Tables.noStyle())))
        .mutate(
            Targets.table("BudgetTable")
                .rowByKey("Item", Values.textFile(Path.of("authored-inputs", "item.txt")))
                .cell("Amount")
                .set(Values.number(125.0)))
        .inspect(
            Targets.table("BudgetTable")
                .rowByKey("Item", Values.textFile(Path.of("authored-inputs", "item.txt")))
                .cell("Amount")
                .read())
        .assertThat(
            Targets.table("BudgetTable")
                .rowByKey("Item", Values.textFile(Path.of("authored-inputs", "item.txt")))
                .cell("Amount")
                .valueEquals(Values.expectedNumber(125.0)));
```

This is still the same GridGrind contract: the Java layer is simply constructing selectors,
actions, queries, and assertions without hand-writing JSON.

If you want to run that example against the repository checkout directly, use `examples/` as the
workspace root so the example can read the committed
[../examples/authored-inputs/item.txt](../examples/authored-inputs/item.txt) input file.

## Emit JSON Or Execute Explicitly

`GridGrindPlan` stays at the canonical contract boundary:

- `toPlan()` returns the immutable `WorkbookPlan`
- `toJsonBytes()`, `toJsonString()`, and `writeJson(...)` emit the same request JSON that the CLI
  would accept

If you want to execute in process, pass that plan to an executor yourself:

Typical in-process execution:

```java
GridGrindResponse response =
    new DefaultGridGrindRequestExecutor()
        .execute(
            plan.toPlan(),
            new ExecutionInputBindings(workspace),
            ExecutionJournalSink.NOOP);
```

## Source-Backed Inputs From Java

The Java layer emits the same source-backed payload contract as JSON requests:

- `Values.textFile(path)`, `Values.formulaFile(path)`, and `Values.binaryFile(path)` emit
  `UTF8_FILE` or `FILE` sources
- `Values.textFromStandardInput()`, `Values.formulaFromStandardInput()`, and
  `Values.binaryFromStandardInput()` emit `STANDARD_INPUT` sources
- relative source-backed paths resolve from `ExecutionInputBindings.workingDirectory()`, not from
  the Java source file location
- `STANDARD_INPUT` values are only usable in-process when the bound
  `ExecutionInputBindings` carries stdin bytes

Use explicit bindings whenever the plan depends on relative authored-input files or caller-supplied
stdin content.

## When To Choose Java Authoring

- You are generating workbook workflows inside an existing JVM application or service.
- You want selector-first plan construction without hand-authoring JSON strings.
- You want exact contract parity with the CLI and JSON protocol while staying in Java.

If you want the artifact-first workflow instead, start with [QUICK_START.md](./QUICK_START.md).
