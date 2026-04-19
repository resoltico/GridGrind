package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction;
import java.util.List;

/** Generated examples that focus on charts, pivots, packages, and binary payload handling. */
final class WorkbookAssetExamples {
  private static final String ONE_PIXEL_PNG_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=";

  private WorkbookAssetExamples() {}

  static GridGrindShippedExamples.ShippedExample sourceBackedInputExample() {
    return ExamplePlanSupport.example(
        "SOURCE_BACKED_INPUT",
        "source-backed-input-request.json",
        "File-backed text, formula, and binary payload authoring without large inline literals.",
        ExamplePlanSupport.plan(
            "source-backed-input-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            null,
            ExamplePlanSupport.step(
                "ensure-inputs",
                ExamplePlanSupport.sheet("Inputs"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "seed-values",
                ExamplePlanSupport.range("Inputs", "B2:B3"),
                new MutationAction.SetRange(
                    ExamplePlanSupport.rows(
                        ExamplePlanSupport.row(ExamplePlanSupport.number(12.0d)),
                        ExamplePlanSupport.row(ExamplePlanSupport.number(18.0d))))),
            ExamplePlanSupport.step(
                "set-title-from-file",
                ExamplePlanSupport.cell("Inputs", "A1"),
                new MutationAction.SetCell(
                    new CellInput.Text(
                        TextSourceInput.utf8File(
                            "examples/source-backed-input-assets/title.txt")))),
            ExamplePlanSupport.step(
                "set-total-formula-from-file",
                ExamplePlanSupport.cell("Inputs", "B4"),
                new MutationAction.SetCell(
                    new CellInput.Formula(
                        TextSourceInput.utf8File(
                            "examples/source-backed-input-assets/total-formula.txt")))),
            ExamplePlanSupport.step(
                "attach-payload-from-file",
                ExamplePlanSupport.sheet("Inputs"),
                new MutationAction.SetEmbeddedObject(
                    new EmbeddedObjectInput(
                        "InputsPayload",
                        "Inputs payload",
                        "inputs-payload.txt",
                        "open",
                        BinarySourceInput.file("examples/source-backed-input-assets/payload.bin"),
                        new PictureDataInput(
                            ExcelPictureFormat.PNG,
                            BinarySourceInput.inlineBase64(ONE_PIXEL_PNG_BASE64)),
                        ExamplePlanSupport.anchor(3, 0, 5, 4)))),
            ExamplePlanSupport.read(
                "read-cells",
                ExamplePlanSupport.cells("Inputs", "A1", "B4"),
                new InspectionQuery.GetCells()),
            ExamplePlanSupport.read(
                "read-payload",
                new DrawingObjectSelector.ByName("Inputs", "InputsPayload"),
                new InspectionQuery.GetDrawingObjectPayload())));
  }

  static GridGrindShippedExamples.ShippedExample chartExample() {
    return ExamplePlanSupport.example(
        "CHART",
        "chart-request.json",
        "Supported simple-chart authoring with named-range-backed series and factual chart readback.",
        ExamplePlanSupport.plan(
            "chart-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs("cli/build/generated-workbooks/gridgrind-chart.xlsx"),
            null,
            ExamplePlanSupport.step(
                "step-01-ensure-sheet",
                ExamplePlanSupport.sheet("Ops"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-set-range",
                ExamplePlanSupport.range("Ops", "A1:C4"),
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
                "step-03-set-categories",
                new NamedRangeSelector.WorkbookScope("ChartCategories"),
                new MutationAction.SetNamedRange(
                    "ChartCategories",
                    new NamedRangeScope.Workbook(),
                    new NamedRangeTarget("Ops", "A2:A4"))),
            ExamplePlanSupport.step(
                "step-04-set-actual",
                new NamedRangeSelector.WorkbookScope("ChartActual"),
                new MutationAction.SetNamedRange(
                    "ChartActual",
                    new NamedRangeScope.Workbook(),
                    new NamedRangeTarget("Ops", "C2:C4"))),
            ExamplePlanSupport.step(
                "step-05-set-chart",
                ExamplePlanSupport.sheet("Ops"),
                new MutationAction.SetChart(
                    new ChartInput.Bar(
                        "OpsChart",
                        ExamplePlanSupport.anchor(4, 0, 8, 12),
                        new ChartInput.Title.Text(TextSourceInput.inline("Roadmap")),
                        new ChartInput.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                        ExcelChartDisplayBlanksAs.SPAN,
                        false,
                        true,
                        ExcelChartBarDirection.COLUMN,
                        List.of(
                            new ChartInput.Series(
                                new ChartInput.Title.Text(TextSourceInput.inline("Plan")),
                                new ChartInput.DataSource("ChartCategories"),
                                new ChartInput.DataSource("Ops!$B$2:$B$4")),
                            new ChartInput.Series(
                                new ChartInput.Title.Text(TextSourceInput.inline("Actual")),
                                new ChartInput.DataSource("ChartCategories"),
                                new ChartInput.DataSource("ChartActual")))))),
            ExamplePlanSupport.read(
                "charts", ExamplePlanSupport.sheet("Ops"), new InspectionQuery.GetCharts()),
            ExamplePlanSupport.assertStep(
                "chart-present",
                new ChartSelector.ByName("Ops", "OpsChart"),
                new Assertion.Present())));
  }

  static GridGrindShippedExamples.ShippedExample pivotExample() {
    return ExamplePlanSupport.example(
        "PIVOT",
        "pivot-request.json",
        "Pivot authoring from a contiguous range with pivot readback and health analysis.",
        ExamplePlanSupport.plan(
            "pivot-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs("cli/build/generated-workbooks/gridgrind-pivot.xlsx"),
            null,
            ExamplePlanSupport.step(
                "step-01-ensure-data",
                ExamplePlanSupport.sheet("Data"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-ensure-report",
                ExamplePlanSupport.sheet("RangeReport"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-03-set-range",
                ExamplePlanSupport.range("Data", "A1:D5"),
                new MutationAction.SetRange(
                    ExamplePlanSupport.rows(
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Region"),
                            ExamplePlanSupport.text("Stage"),
                            ExamplePlanSupport.text("Owner"),
                            ExamplePlanSupport.text("Amount")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("North"),
                            ExamplePlanSupport.text("Plan"),
                            ExamplePlanSupport.text("Ada"),
                            ExamplePlanSupport.number(10.0d)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("North"),
                            ExamplePlanSupport.text("Do"),
                            ExamplePlanSupport.text("Ada"),
                            ExamplePlanSupport.number(15.0d)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("South"),
                            ExamplePlanSupport.text("Plan"),
                            ExamplePlanSupport.text("Lin"),
                            ExamplePlanSupport.number(7.0d)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("South"),
                            ExamplePlanSupport.text("Do"),
                            ExamplePlanSupport.text("Lin"),
                            ExamplePlanSupport.number(12.0d))))),
            ExamplePlanSupport.step(
                "step-04-set-pivot",
                new PivotTableSelector.ByNameOnSheet("RegionalTotals", "RangeReport"),
                new MutationAction.SetPivotTable(
                    new PivotTableInput(
                        "RegionalTotals",
                        "RangeReport",
                        new PivotTableInput.Source.Range("Data", "A1:D5"),
                        new PivotTableInput.Anchor("A3"),
                        List.of("Region"),
                        List.of("Stage"),
                        List.of(),
                        List.of(
                            new PivotTableInput.DataField(
                                "Amount",
                                ExcelPivotDataConsolidateFunction.SUM,
                                "Total Amount",
                                "#,##0.00"))))),
            ExamplePlanSupport.read(
                "pivot-tables",
                new PivotTableSelector.ByName("RegionalTotals"),
                new InspectionQuery.GetPivotTables()),
            ExamplePlanSupport.read(
                "pivot-health",
                new PivotTableSelector.ByName("RegionalTotals"),
                new InspectionQuery.AnalyzePivotTableHealth())));
  }

  static GridGrindShippedExamples.ShippedExample packageSecurityInspectionExample() {
    return ExamplePlanSupport.example(
        "PACKAGE_SECURITY_INSPECTION",
        "package-security-inspect-request.json",
        "Encrypted package open plus factual package-security and cell inspection.",
        ExamplePlanSupport.plan(
            "package-security-inspection-workflow",
            new WorkbookPlan.WorkbookSource.ExistingFile(
                "cli/build/generated-workbooks/gridgrind-package-security.xlsx",
                new OoxmlOpenSecurityInput("GridGrind-2026")),
            new WorkbookPlan.WorkbookPersistence.None(),
            null,
            ExamplePlanSupport.read(
                "security",
                ExamplePlanSupport.workbook(),
                new InspectionQuery.GetPackageSecurity()),
            ExamplePlanSupport.read(
                "secure-cells",
                ExamplePlanSupport.cells("Secure", "A1", "A2", "B2", "A3", "B3"),
                new InspectionQuery.GetCells())));
  }
}
