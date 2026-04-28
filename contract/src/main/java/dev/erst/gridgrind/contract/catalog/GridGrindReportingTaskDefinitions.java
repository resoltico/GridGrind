package dev.erst.gridgrind.contract.catalog;

import static dev.erst.gridgrind.contract.catalog.GridGrindTaskDefinitionSupport.definition;
import static dev.erst.gridgrind.contract.catalog.GridGrindTaskDefinitionSupport.phase;
import static dev.erst.gridgrind.contract.catalog.GridGrindTaskDefinitionSupport.protocolLookupNote;
import static dev.erst.gridgrind.contract.catalog.GridGrindTaskDefinitionSupport.rebasePlan;
import static dev.erst.gridgrind.contract.catalog.GridGrindTaskDefinitionSupport.ref;
import static dev.erst.gridgrind.contract.catalog.GridGrindTaskDefinitionSupport.task;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import java.util.List;

/** Contract-owned task definitions for report, dashboard, intake, and pivot workflows. */
final class GridGrindReportingTaskDefinitions {
  private GridGrindReportingTaskDefinitions() {}

  static List<TaskDefinition> definitions() {
    return List.of(tabularReport(), dashboard(), dataEntryWorkflow(), pivotReport());
  }

  private static TaskDefinition tabularReport() {
    TaskEntry task =
        task(
            "TABULAR_REPORT",
            "Create a structured worksheet report with typed cells, table semantics, and factual"
                + " readback.",
            List.of("office", "reporting", "table", "report", "save"),
            List.of(
                "Sheet structure is created intentionally instead of ad hoc cell drift.",
                "Rows can be modeled as a table so filtering and later readback stay"
                    + " authoritative.",
                "Critical facts can be inspected before the workbook is persisted."),
            List.of(
                "Sheet name and header layout.",
                "Typed row values or source-backed payload files for larger authored content.",
                "Persistence target when the result must be saved."),
            List.of(
                "Totals rows and formula-backed summaries.",
                "Cell styling and print-layout refinements.",
                "Assertions on key balances or counts."),
            List.of(
                phase(
                    "Lay Out The Workbook",
                    "Create sheets, headers, and fixed structure before data rows arrive.",
                    List.of(
                        ref("sourceTypes", "NEW"),
                        ref("mutationActionTypes", "ENSURE_SHEET"),
                        ref("mutationActionTypes", "SET_RANGE")),
                    List.of(
                        "Use one intentional sheet skeleton instead of scattered cell writes.")),
                phase(
                    "Model The Table",
                    "Move the sheet from loose cells into tabular semantics.",
                    List.of(
                        ref("mutationActionTypes", "SET_TABLE"),
                        ref("mutationActionTypes", "AUTO_SIZE_COLUMNS"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("Prefer one table per logical data region.")),
                phase(
                    "Verify And Inspect",
                    "Read back the cells or workbook facts that make the report trustworthy.",
                    List.of(
                        ref("inspectionQueryTypes", "GET_CELLS"),
                        ref("inspectionQueryTypes", "GET_TABLES"),
                        ref("inspectionQueryTypes", "GET_WORKBOOK_SUMMARY")),
                    List.of("Use factual readback instead of assuming writes landed correctly."))),
            List.of(
                "Large authored literals belong in UTF8_FILE, FILE, or STANDARD_INPUT sources"
                    + " instead of huge inline JSON.",
                "Table headers must remain nonblank and unique.",
                "Formula authoring is scalar-only unless you use SET_ARRAY_FORMULA."));
    return definition(
        task,
        List.of(
            "tabular report",
            "structured worksheet",
            "build a report table",
            "typed worksheet report",
            "create a report workbook"),
        ExamplePlanSupport.plan(
            "tabular-report-starter",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs("tabular-report.xlsx"),
            null,
            ExamplePlanSupport.step(
                "ensure-report",
                ExamplePlanSupport.sheet("Report"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "seed-report-rows",
                ExamplePlanSupport.range("Report", "A1:C4"),
                new MutationAction.SetRange(
                    ExamplePlanSupport.rows(
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Item"),
                            ExamplePlanSupport.text("Amount"),
                            ExamplePlanSupport.text("Status")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Hosting"),
                            ExamplePlanSupport.number(49.0d),
                            ExamplePlanSupport.text("APPROVED")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Domain"),
                            ExamplePlanSupport.number(12.0d),
                            ExamplePlanSupport.text("APPROVED")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Monitoring"),
                            ExamplePlanSupport.number(25.0d),
                            ExamplePlanSupport.text("REVIEW"))))),
            ExamplePlanSupport.step(
                "define-report-table",
                ExamplePlanSupport.table("ReportTable", "Report"),
                new MutationAction.SetTable(
                    new TableInput(
                        "ReportTable",
                        "Report",
                        "A1:C4",
                        false,
                        new TableStyleInput.Named(
                            "TableStyleMedium2", false, false, true, false)))),
            ExamplePlanSupport.step(
                "auto-size",
                ExamplePlanSupport.sheet("Report"),
                new MutationAction.AutoSizeColumns()),
            ExamplePlanSupport.read(
                "report-cells",
                ExamplePlanSupport.cells("Report", "A1", "B4", "C4"),
                new InspectionQuery.GetCells()),
            ExamplePlanSupport.read(
                "tables",
                new TableSelector.ByNameOnSheet("ReportTable", "Report"),
                new InspectionQuery.GetTables())),
        List.of(
            "The starter plan is runnable as-is and gives you one typed report table to adapt.",
            "Replace the sample headers, rows, and output filename with your own workbook facts.",
            protocolLookupNote(
                "table and row-ingestion request shapes",
                List.of("mutationActionTypes:SET_TABLE", "mutationActionTypes:APPEND_ROW"),
                List.of("table", "append row"))));
  }

  private static TaskDefinition dashboard() {
    TaskEntry task =
        task(
            "DASHBOARD",
            "Assemble an executive dashboard from reusable named surfaces and supported charts.",
            List.of("office", "dashboard", "charts", "summary", "kpi"),
            List.of(
                "Summary sheets and KPI surfaces are intentionally structured.",
                "Reusable named surfaces back formulas or charts instead of fragile copied"
                    + " ranges.",
                "Charts are authored and then read back through the same contract surface."),
            List.of(
                "Metric definitions and source ranges.",
                "Chart names and anchors.",
                "Target persistence path when the dashboard must be saved."),
            List.of(
                "Assertions that required dashboard entities exist.",
                "Workbook-health analysis after authoring.",
                "Named-range-backed chart series for reusable models."),
            List.of(
                phase(
                    "Assemble Summary Sheets",
                    "Create the dashboard canvas and key text or formula cells first.",
                    List.of(
                        ref("sourceTypes", "NEW"),
                        ref("mutationActionTypes", "ENSURE_SHEET"),
                        ref("mutationActionTypes", "SET_CELL"),
                        ref("mutationActionTypes", "SET_RANGE")),
                    List.of("Keep summary layout intentional so later chart anchors are stable.")),
                phase(
                    "Define Reusable Model Surfaces",
                    "Create named ranges that charts and formulas can depend on.",
                    List.of(
                        ref("mutationActionTypes", "SET_NAMED_RANGE"),
                        ref("mutationActionTypes", "SET_CHART")),
                    List.of("Named surfaces reduce accidental drift when the dashboard evolves.")),
                phase(
                    "Author And Inspect Visuals",
                    "Create supported charts and verify that the expected entities exist.",
                    List.of(
                        ref("inspectionQueryTypes", "GET_CHARTS"),
                        ref("assertionTypes", "EXPECT_CHART_PRESENT"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of(
                        "Use factual chart readback to confirm what the workbook now contains."))),
            List.of(
                "SET_CHART supports the authoritative simple-chart family listed in the protocol"
                    + " catalog.",
                "Chart title and series FORMULA titles must resolve to one cell.",
                "Unsupported loaded chart detail is preserved on unrelated edits but is not"
                    + " available for authoritative mutation."));
    return definition(
        task,
        List.of(
            "dashboard",
            "kpi summary",
            "executive dashboard",
            "monthly sales dashboard",
            "chart workbook"),
        ExamplePlanSupport.plan(
            "dashboard-starter",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs("dashboard-output.xlsx"),
            null,
            ExamplePlanSupport.step(
                "ensure-data", ExamplePlanSupport.sheet("Data"), new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "seed-data",
                ExamplePlanSupport.range("Data", "A1:C4"),
                new MutationAction.SetRange(
                    ExamplePlanSupport.rows(
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Month"),
                            ExamplePlanSupport.text("Plan"),
                            ExamplePlanSupport.text("Actual")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Jan"),
                            ExamplePlanSupport.number(10.0d),
                            ExamplePlanSupport.number(12.0d)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Feb"),
                            ExamplePlanSupport.number(18.0d),
                            ExamplePlanSupport.number(17.0d)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Mar"),
                            ExamplePlanSupport.number(15.0d),
                            ExamplePlanSupport.number(16.0d))))),
            ExamplePlanSupport.step(
                "ensure-dashboard",
                ExamplePlanSupport.sheet("Dashboard"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "set-title",
                ExamplePlanSupport.cell("Dashboard", "A1"),
                new MutationAction.SetCell(ExamplePlanSupport.text("Monthly Sales Dashboard"))),
            ExamplePlanSupport.step(
                "set-categories",
                new NamedRangeSelector.WorkbookScope("DashboardCategories"),
                new MutationAction.SetNamedRange(
                    "DashboardCategories",
                    new NamedRangeScope.Workbook(),
                    new NamedRangeTarget("Data", "A2:A4"))),
            ExamplePlanSupport.step(
                "set-actual",
                new NamedRangeSelector.WorkbookScope("DashboardActual"),
                new MutationAction.SetNamedRange(
                    "DashboardActual",
                    new NamedRangeScope.Workbook(),
                    new NamedRangeTarget("Data", "C2:C4"))),
            ExamplePlanSupport.step(
                "set-chart",
                ExamplePlanSupport.sheet("Dashboard"),
                new MutationAction.SetChart(
                    new ChartInput(
                        "RevenueChart",
                        ExamplePlanSupport.anchor(4, 1, 9, 15),
                        new ChartInput.Title.Text(TextSourceInput.inline("Revenue Overview")),
                        new ChartInput.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                        ExcelChartDisplayBlanksAs.SPAN,
                        false,
                        List.of(
                            new ChartInput.Bar(
                                true,
                                ExcelChartBarDirection.COLUMN,
                                null,
                                null,
                                null,
                                null,
                                List.of(
                                    new ChartInput.Series(
                                        new ChartInput.Title.Text(TextSourceInput.inline("Plan")),
                                        new ChartInput.DataSource.Reference("DashboardCategories"),
                                        new ChartInput.DataSource.Reference("Data!$B$2:$B$4"),
                                        null,
                                        null,
                                        null,
                                        null),
                                    new ChartInput.Series(
                                        new ChartInput.Title.Text(TextSourceInput.inline("Actual")),
                                        new ChartInput.DataSource.Reference("DashboardCategories"),
                                        new ChartInput.DataSource.Reference("DashboardActual"),
                                        null,
                                        null,
                                        null,
                                        null))))))),
            ExamplePlanSupport.read(
                "charts",
                new ChartSelector.AllOnSheet("Dashboard"),
                new InspectionQuery.GetCharts())),
        List.of(
            "The starter plan is runnable and seeds one simple dashboard with named-range-backed"
                + " chart series.",
            "Swap the sample data, dashboard title, and chart family for your own KPI model.",
            protocolLookupNote(
                "chart and reusable-model request shapes",
                List.of("mutationActionTypes:SET_CHART", "mutationActionTypes:SET_NAMED_RANGE"),
                List.of("chart", "named range"))));
  }

  private static TaskDefinition dataEntryWorkflow() {
    TaskEntry task =
        task(
            "DATA_ENTRY_WORKFLOW",
            "Build an intake-style sheet, attach validation, then append rows and spot-check the"
                + " result.",
            List.of("office", "data-entry", "validation", "append", "workflow"),
            List.of(
                "One intake sheet is structured before row ingestion begins.",
                "Business rules can be expressed through validation and typed writes.",
                "Key fields are read back before the workbook is saved."),
            List.of(
                "Schema or header layout.",
                "Row payload source.",
                "Persistence target when the workbook should leave memory."),
            List.of(
                "Append-only large-file plans.",
                "Validation rules on business-critical columns.",
                "Post-write factual readback for spot checks."),
            List.of(
                phase(
                    "Prepare The Intake Surface",
                    "Create the intake sheet and attach the first layer of operator guidance.",
                    List.of(
                        ref("sourceTypes", "NEW"),
                        ref("mutationActionTypes", "ENSURE_SHEET"),
                        ref("mutationActionTypes", "APPEND_ROW"),
                        ref("mutationActionTypes", "SET_DATA_VALIDATION")),
                    List.of("Stabilize the schema before appending many rows.")),
                phase(
                    "Load Rows",
                    "Append data in a way that keeps the workflow readable and deterministic.",
                    List.of(
                        ref("mutationActionTypes", "APPEND_ROW"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("For large append-only flows, evaluate whether STREAMING_WRITE fits.")),
                phase(
                    "Spot-Check The Result",
                    "Read back critical cells instead of trusting write success alone.",
                    List.of(ref("inspectionQueryTypes", "GET_CELLS")),
                    List.of("Read back the fields that operators would inspect manually."))),
            List.of(
                "STREAMING_WRITE is append-only and requires a NEW source with constrained"
                    + " mutation support.",
                "STANDARD_INPUT-authored values require --request because stdin cannot carry both"
                    + " the request JSON and authored payload content.",
                "Validation rules guide operators, but they do not replace factual readback."));
    return definition(
        task,
        List.of(
            "data entry workflow",
            "intake sheet",
            "validated intake workbook",
            "append rows with validation",
            "operator workbook"),
        ExamplePlanSupport.plan(
            "data-entry-starter",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs("data-entry-output.xlsx"),
            ExecutionPolicyInput.defaults(),
            ExamplePlanSupport.step(
                "ensure-intake",
                ExamplePlanSupport.sheet("Intake"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "append-header",
                ExamplePlanSupport.sheet("Intake"),
                new MutationAction.AppendRow(
                    List.of(
                        ExamplePlanSupport.text("Ticket"),
                        ExamplePlanSupport.text("Owner"),
                        ExamplePlanSupport.text("Status")))),
            ExamplePlanSupport.step(
                "set-status-validation",
                ExamplePlanSupport.range("Intake", "C2:C500"),
                new MutationAction.SetDataValidation(
                    new DataValidationInput(
                        new DataValidationRuleInput.ExplicitList(List.of("NEW", "REVIEW", "DONE")),
                        true,
                        null,
                        null,
                        null))),
            ExamplePlanSupport.step(
                "append-sample-row",
                ExamplePlanSupport.sheet("Intake"),
                new MutationAction.AppendRow(
                    List.of(
                        ExamplePlanSupport.text("T-100"),
                        ExamplePlanSupport.text("Ada"),
                        ExamplePlanSupport.text("NEW")))),
            ExamplePlanSupport.read(
                "spot-check",
                ExamplePlanSupport.cells("Intake", "A2", "B2", "C2"),
                new InspectionQuery.GetCells())),
        List.of(
            "The starter plan is runnable and shows header creation, validation, one appended row,"
                + " and factual reread.",
            "Replace the sample schema, validation list, and output filename with your real"
                + " intake workflow.",
            protocolLookupNote(
                "validation and large-file execution shapes",
                List.of(
                    "mutationActionTypes:SET_DATA_VALIDATION", "plainTypes:executionModeInputType"),
                List.of("data validation", "streaming write"))));
  }

  private static TaskDefinition pivotReport() {
    TaskEntry task =
        task(
            "PIVOT_REPORT",
            "Build a pivot-backed report or inspect pivot metadata and health on supported pivot"
                + " tables.",
            List.of("office", "pivot", "summary", "analysis", "report"),
            List.of(
                "Pivot-backed summaries are authored from contiguous ranges, named ranges, or"
                    + " tables.",
                "Pivot metadata is surfaced back as facts instead of screenshots or guesses.",
                "Pivot health can be analyzed after authoring or against loaded workbooks."),
            List.of(
                "Pivot source data and column roles.",
                "Destination sheet and anchor.",
                "Persistence target when the workbook should be saved."),
            List.of(
                "Multiple pivot reports against one prepared source table.",
                "Named-range-backed or table-backed pivot sources.",
                "Pivot-health analysis after save or reopen."),
            List.of(
                phase(
                    "Prepare The Source",
                    "Create the source sheet and contiguous data region first.",
                    List.of(
                        ref("sourceTypes", "NEW"),
                        ref("mutationActionTypes", "ENSURE_SHEET"),
                        ref("mutationActionTypes", "SET_RANGE")),
                    List.of("Pivot authoring depends on a prepared source surface.")),
                phase(
                    "Author The Pivot",
                    "Create one workbook-global pivot-table definition on the report sheet.",
                    List.of(
                        ref("mutationActionTypes", "SET_PIVOT_TABLE"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("Keep row, column, report-filter, and data-field roles disjoint.")),
                phase(
                    "Inspect The Result",
                    "Read back pivot metadata and analyze pivot health after authoring.",
                    List.of(
                        ref("inspectionQueryTypes", "GET_PIVOT_TABLES"),
                        ref("inspectionQueryTypes", "ANALYZE_PIVOT_TABLE_HEALTH")),
                    List.of("Pivot health explains malformed or unsupported loaded detail."))),
            List.of(
                "Pivot authoring is limited to the supported XSSF surface documented publicly.",
                "Loaded unsupported pivot detail is surfaced explicitly instead of rewritten.",
                "Report filters impose anchor constraints documented in the pivot mutation"
                    + " reference."));
    return definition(
        task,
        List.of(
            "pivot report",
            "pivot table workbook",
            "pivot summary",
            "build pivots",
            "pivot health"),
        rebasePlan(
            WorkbookAssetExamples.pivotExample(ExamplePathLayout.BUILT_IN).plan(),
            "pivot-report-starter",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs("pivot-report.xlsx")),
        List.of(
            "The starter plan is runnable and demonstrates seeded source data, authored pivot"
                + " metadata, and pivot-health analysis.",
            "Replace the sample dimensions, aggregation, and output filename with your own source"
                + " layout.",
            protocolLookupNote(
                "pivot authoring and inspection shapes",
                List.of(
                    "mutationActionTypes:SET_PIVOT_TABLE",
                    "inspectionQueryTypes:ANALYZE_PIVOT_TABLE_HEALTH"),
                List.of("pivot"))));
  }
}
