package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.util.Base64;
import java.util.List;

/** Representative workbook command and read-command samples for exhaustive coverage tests. */
final class WorkbookSampleFixtures {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  private WorkbookSampleFixtures() {}

  static List<WorkbookReadCommand.Introspection> introspectionCommands() {
    return List.of(
        new WorkbookReadCommand.GetWorkbookSummary("workbook"),
        new WorkbookReadCommand.GetWorkbookProtection("protection"),
        new WorkbookReadCommand.GetCustomXmlMappings("customXmlMappings"),
        new WorkbookReadCommand.ExportCustomXmlMapping(
            "customXmlExport",
            new ExcelCustomXmlMappingLocator(1L, "CORSO_mapping"),
            false,
            StandardCharsets.UTF_8.name()),
        new WorkbookReadCommand.GetNamedRanges("ranges", new ExcelNamedRangeSelection.All()),
        new WorkbookReadCommand.GetSheetSummary("sheet", "Budget"),
        new WorkbookReadCommand.GetCells("cells", "Budget", List.of("A1")),
        new WorkbookReadCommand.GetWindow("window", "Budget", "A1", 1, 1),
        new WorkbookReadCommand.GetMergedRegions("merged", "Budget"),
        new WorkbookReadCommand.GetHyperlinks(
            "hyperlinks", "Budget", new ExcelCellSelection.AllUsedCells()),
        new WorkbookReadCommand.GetComments(
            "comments", "Budget", new ExcelCellSelection.Selected(List.of("A1"))),
        new WorkbookReadCommand.GetDrawingObjects("drawings", "Budget"),
        new WorkbookReadCommand.GetCharts("charts", "Budget"),
        new WorkbookReadCommand.GetPivotTables("pivots", new ExcelPivotTableSelection.All()),
        new WorkbookReadCommand.GetDrawingObjectPayload("payload", "Budget", "Object 1"),
        new WorkbookReadCommand.GetSheetLayout("layout", "Budget"),
        new WorkbookReadCommand.GetPrintLayout("print", "Budget"),
        new WorkbookReadCommand.GetDataValidations(
            "validations", "Budget", new ExcelRangeSelection.All()),
        new WorkbookReadCommand.GetConditionalFormatting(
            "conditional", "Budget", new ExcelRangeSelection.Selected(List.of("A1:B2"))),
        new WorkbookReadCommand.GetAutofilters("autofilters", "Budget"),
        new WorkbookReadCommand.GetTables("tables", new ExcelTableSelection.All()),
        new WorkbookReadCommand.GetArrayFormulas(
            "arrayFormulas", new ExcelSheetSelection.Selected(List.of("Budget"))),
        new WorkbookReadCommand.GetFormulaSurface(
            "formula", new ExcelSheetSelection.Selected(List.of("Budget"))),
        new WorkbookReadCommand.GetSheetSchema("schema", "Budget", "A1", 1, 1),
        new WorkbookReadCommand.GetNamedRangeSurface(
            "namedSurface", new ExcelNamedRangeSelection.All()));
  }

  static List<WorkbookCommand> workbookCommands() {
    ExcelDrawingAnchor.TwoCell firstAnchor = anchor(1, 1, 4, 6);
    ExcelDrawingAnchor.TwoCell movedAnchor =
        new ExcelDrawingAnchor.TwoCell(
            new ExcelDrawingMarker(6, 2),
            new ExcelDrawingMarker(9, 7),
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);

    return List.of(
        new WorkbookSheetCommand.CreateSheet("Budget"),
        new WorkbookSheetCommand.RenameSheet("Budget", "Summary"),
        new WorkbookSheetCommand.DeleteSheet("Archive"),
        new WorkbookSheetCommand.MoveSheet("Budget", 1),
        new WorkbookSheetCommand.CopySheet(
            "Budget", "Budget Copy", new ExcelSheetCopyPosition.AppendAtEnd()),
        new WorkbookSheetCommand.SetActiveSheet("Budget"),
        new WorkbookSheetCommand.SetSelectedSheets(List.of("Budget", "Summary")),
        new WorkbookSheetCommand.SetSheetVisibility("Budget", ExcelSheetVisibility.VERY_HIDDEN),
        new WorkbookSheetCommand.SetSheetProtection("Budget", protectionSettings()),
        new WorkbookSheetCommand.ClearSheetProtection("Budget"),
        new WorkbookSheetCommand.SetWorkbookProtection(
            new ExcelWorkbookProtectionSettings(true, false, true, "book", "review")),
        new WorkbookSheetCommand.ClearWorkbookProtection(),
        new WorkbookStructureCommand.MergeCells("Budget", "A1:B2"),
        new WorkbookStructureCommand.UnmergeCells("Budget", "A1:B2"),
        new WorkbookStructureCommand.SetColumnWidth("Budget", 0, 1, 16.0d),
        new WorkbookStructureCommand.SetRowHeight("Budget", 0, 2, 28.5d),
        new WorkbookStructureCommand.InsertRows("Budget", 2, 3),
        new WorkbookStructureCommand.DeleteRows("Budget", new ExcelRowSpan(4, 6)),
        new WorkbookStructureCommand.ShiftRows("Budget", new ExcelRowSpan(1, 3), 2),
        new WorkbookStructureCommand.InsertColumns("Budget", 1, 2),
        new WorkbookStructureCommand.DeleteColumns("Budget", new ExcelColumnSpan(3, 4)),
        new WorkbookStructureCommand.ShiftColumns("Budget", new ExcelColumnSpan(0, 1), -1),
        new WorkbookStructureCommand.SetRowVisibility("Budget", new ExcelRowSpan(5, 7), true),
        new WorkbookStructureCommand.SetColumnVisibility(
            "Budget", new ExcelColumnSpan(2, 3), false),
        new WorkbookStructureCommand.GroupRows("Budget", new ExcelRowSpan(8, 10), false),
        new WorkbookStructureCommand.UngroupRows("Budget", new ExcelRowSpan(8, 10)),
        new WorkbookStructureCommand.GroupColumns("Budget", new ExcelColumnSpan(4, 6), true),
        new WorkbookStructureCommand.UngroupColumns("Budget", new ExcelColumnSpan(4, 6)),
        new WorkbookLayoutCommand.SetSheetPane("Budget", new ExcelSheetPane.Frozen(1, 2, 1, 2)),
        new WorkbookLayoutCommand.SetSheetZoom("Budget", 135),
        new WorkbookLayoutCommand.SetPrintLayout("Budget", defaultPrintLayout()),
        new WorkbookLayoutCommand.ClearPrintLayout("Budget"),
        new WorkbookCellCommand.SetCell(
            "Budget", "A1", ExcelCellValue.date(LocalDate.of(2026, 3, 23))),
        new WorkbookCellCommand.SetArrayFormula(
            "Budget", "D2:D4", new ExcelArrayFormulaDefinition("B2:B4*C2:C4")),
        new WorkbookCellCommand.ClearArrayFormula("Budget", "D2"),
        new WorkbookMetadataCommand.ImportCustomXmlMapping(
            new ExcelCustomXmlImportDefinition(
                new ExcelCustomXmlMappingLocator(1L, "CORSO_mapping"),
                "<CORSO><NOME>Ops</NOME></CORSO>")),
        new WorkbookCellCommand.SetRange(
            "Budget",
            "A1:B2",
            List.of(
                List.of(ExcelCellValue.text("Item"), ExcelCellValue.number(49.0d)),
                List.of(ExcelCellValue.text("Tax"), ExcelCellValue.number(10.0d)))),
        new WorkbookCellCommand.ClearRange("Budget", "C1:C2"),
        new WorkbookAnnotationCommand.SetHyperlink(
            "Budget", "A1", new ExcelHyperlink.Url("https://example.com/report")),
        new WorkbookAnnotationCommand.ClearHyperlink("Budget", "A1"),
        new WorkbookAnnotationCommand.SetComment(
            "Budget", "A1", new ExcelComment("Review", "GridGrind", false)),
        new WorkbookAnnotationCommand.ClearComment("Budget", "A1"),
        new WorkbookDrawingCommand.SetPicture(
            "Budget",
            new ExcelPictureDefinition(
                "BudgetPicture",
                new ExcelBinaryData(PNG_PIXEL_BYTES),
                ExcelPictureFormat.PNG,
                firstAnchor,
                "Queue preview")),
        new WorkbookDrawingCommand.SetSignatureLine(
            "Budget",
            new ExcelSignatureLineDefinition(
                "BudgetSignature",
                firstAnchor,
                false,
                "Review the budget before signing.",
                "Ada Lovelace",
                "Finance",
                "ada@example.com",
                null,
                "invalid",
                ExcelPictureFormat.PNG,
                new ExcelBinaryData(PNG_PIXEL_BYTES))),
        new WorkbookDrawingCommand.SetChart("Budget", chartDefinition(firstAnchor)),
        new WorkbookTabularCommand.SetPivotTable(
            new ExcelPivotTableDefinition(
                "Budget Pivot",
                "Budget",
                new ExcelPivotTableDefinition.Source.Range("Budget", "A1:C4"),
                new ExcelPivotTableDefinition.Anchor("E3"),
                List.of("Item"),
                List.of(),
                List.of(),
                List.of(
                    new ExcelPivotTableDefinition.DataField(
                        "Tax", ExcelPivotDataConsolidateFunction.SUM, "Total Tax", "#,##0.00")))),
        new WorkbookDrawingCommand.SetShape(
            "Budget",
            new ExcelShapeDefinition(
                "BudgetShape",
                ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                firstAnchor,
                "rect",
                "Queue")),
        new WorkbookDrawingCommand.SetEmbeddedObject(
            "Budget",
            new ExcelEmbeddedObjectDefinition(
                "BudgetEmbed",
                "Payload",
                "payload.txt",
                "payload.txt",
                new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
                ExcelPictureFormat.PNG,
                new ExcelBinaryData(PNG_PIXEL_BYTES),
                firstAnchor)),
        new WorkbookDrawingCommand.SetDrawingObjectAnchor("Budget", "BudgetShape", movedAnchor),
        new WorkbookDrawingCommand.DeleteDrawingObject("Budget", "BudgetPicture"),
        new WorkbookFormattingCommand.ApplyStyle("Budget", "A1:B1", defaultStyle()),
        new WorkbookFormattingCommand.SetDataValidation("Budget", "B2:B5", validationDefinition()),
        new WorkbookFormattingCommand.ClearDataValidations(
            "Budget", new ExcelRangeSelection.Selected(List.of("C2:D4"))),
        new WorkbookFormattingCommand.SetConditionalFormatting(
            "Budget",
            new ExcelConditionalFormattingBlockDefinition(
                List.of("A2:A5"),
                List.of(
                    new ExcelConditionalFormattingRule.FormulaRule(
                        "A2>0",
                        true,
                        new ExcelDifferentialStyle(
                            "0.00", null, null, null, null, null, null, null, null))))),
        new WorkbookFormattingCommand.ClearConditionalFormatting(
            "Budget", new ExcelRangeSelection.Selected(List.of("A2:A5"))),
        new WorkbookTabularCommand.SetAutofilter(
            "Budget",
            "A1:C4",
            List.of(
                new ExcelAutofilterFilterColumn(
                    0L, false, new ExcelAutofilterFilterCriterion.Values(List.of("Ada"), true))),
            new ExcelAutofilterSortState(
                "A1:C4",
                true,
                false,
                java.util.Optional.empty(),
                List.of(new ExcelAutofilterSortCondition.Value("A2:A4", false)))),
        new WorkbookTabularCommand.ClearAutofilter("Budget"),
        new WorkbookTabularCommand.SetTable(
            new ExcelTableDefinition(
                "BudgetTable",
                "Budget",
                "A1:C4",
                true,
                new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false))),
        new WorkbookTabularCommand.DeleteTable("BudgetTable", "Budget"),
        new WorkbookTabularCommand.DeletePivotTable("Budget Pivot", "Budget"),
        new WorkbookMetadataCommand.SetNamedRange(
            new ExcelNamedRangeDefinition(
                "BudgetTotal",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Budget", "B4"))),
        new WorkbookMetadataCommand.DeleteNamedRange(
            "BudgetTotal", new ExcelNamedRangeScope.SheetScope("Budget")),
        new WorkbookCellCommand.AppendRow("Budget", List.of(ExcelCellValue.text("Item"))),
        new WorkbookLayoutCommand.AutoSizeColumns("Budget"));
  }

  static ExcelDrawingAnchor.TwoCell anchor(int fromRow, int fromColumn, int toRow, int toColumn) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromRow, fromColumn),
        new ExcelDrawingMarker(toRow, toColumn),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static ExcelChartDefinition chartDefinition(ExcelDrawingAnchor.TwoCell anchor) {
    return ExcelChartTestSupport.barChart(
        "BudgetChart",
        anchor,
        new ExcelChartDefinition.Title.Text("Roadmap"),
        new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        true,
        ExcelChartBarDirection.COLUMN,
        List.of(
            new ExcelChartDefinition.Series(
                new ExcelChartDefinition.Title.Formula("B1"),
                ExcelChartTestSupport.ref("A2:A4"),
                ExcelChartTestSupport.ref("B2:B4"))));
  }

  private static ExcelDataValidationDefinition validationDefinition() {
    return new ExcelDataValidationDefinition(
        new ExcelDataValidationRule.TextLength(ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
        true,
        false,
        new ExcelDataValidationPrompt("Reason", "Use 20 characters or fewer.", true),
        new ExcelDataValidationErrorAlert(
            ExcelDataValidationErrorStyle.STOP, "Too long", "Use a shorter reason.", true));
  }

  private static ExcelCellStyle defaultStyle() {
    return new ExcelCellStyle(
        "#,##0.00",
        new ExcelCellAlignment(
            true, ExcelHorizontalAlignment.RIGHT, ExcelVerticalAlignment.CENTER, null, null),
        new ExcelCellFont(true, false, null, null, null, null, null),
        null,
        null,
        null);
  }

  private static ExcelPrintLayout defaultPrintLayout() {
    return new ExcelPrintLayout(
        new ExcelPrintLayout.Area.Range("A1:C20"),
        ExcelPrintOrientation.LANDSCAPE,
        new ExcelPrintLayout.Scaling.Fit(1, 0),
        new ExcelPrintLayout.TitleRows.Band(0, 1),
        new ExcelPrintLayout.TitleColumns.Band(0, 0),
        new ExcelHeaderFooterText("Left", "Center", "Right"),
        new ExcelHeaderFooterText("Footer Left", "", "Footer Right"));
  }

  private static ExcelSheetProtectionSettings protectionSettings() {
    return new ExcelSheetProtectionSettings(
        true, false, true, false, true, false, true, false, true, false, true, false, true, false,
        true);
  }
}
