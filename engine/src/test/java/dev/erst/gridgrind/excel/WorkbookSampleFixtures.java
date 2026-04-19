package dev.erst.gridgrind.excel;

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
        new WorkbookCommand.CreateSheet("Budget"),
        new WorkbookCommand.RenameSheet("Budget", "Summary"),
        new WorkbookCommand.DeleteSheet("Archive"),
        new WorkbookCommand.MoveSheet("Budget", 1),
        new WorkbookCommand.CopySheet(
            "Budget", "Budget Copy", new ExcelSheetCopyPosition.AppendAtEnd()),
        new WorkbookCommand.SetActiveSheet("Budget"),
        new WorkbookCommand.SetSelectedSheets(List.of("Budget", "Summary")),
        new WorkbookCommand.SetSheetVisibility("Budget", ExcelSheetVisibility.VERY_HIDDEN),
        new WorkbookCommand.SetSheetProtection("Budget", protectionSettings()),
        new WorkbookCommand.ClearSheetProtection("Budget"),
        new WorkbookCommand.SetWorkbookProtection(
            new ExcelWorkbookProtectionSettings(true, false, true, "book", "review")),
        new WorkbookCommand.ClearWorkbookProtection(),
        new WorkbookCommand.MergeCells("Budget", "A1:B2"),
        new WorkbookCommand.UnmergeCells("Budget", "A1:B2"),
        new WorkbookCommand.SetColumnWidth("Budget", 0, 1, 16.0d),
        new WorkbookCommand.SetRowHeight("Budget", 0, 2, 28.5d),
        new WorkbookCommand.InsertRows("Budget", 2, 3),
        new WorkbookCommand.DeleteRows("Budget", new ExcelRowSpan(4, 6)),
        new WorkbookCommand.ShiftRows("Budget", new ExcelRowSpan(1, 3), 2),
        new WorkbookCommand.InsertColumns("Budget", 1, 2),
        new WorkbookCommand.DeleteColumns("Budget", new ExcelColumnSpan(3, 4)),
        new WorkbookCommand.ShiftColumns("Budget", new ExcelColumnSpan(0, 1), -1),
        new WorkbookCommand.SetRowVisibility("Budget", new ExcelRowSpan(5, 7), true),
        new WorkbookCommand.SetColumnVisibility("Budget", new ExcelColumnSpan(2, 3), false),
        new WorkbookCommand.GroupRows("Budget", new ExcelRowSpan(8, 10), false),
        new WorkbookCommand.UngroupRows("Budget", new ExcelRowSpan(8, 10)),
        new WorkbookCommand.GroupColumns("Budget", new ExcelColumnSpan(4, 6), true),
        new WorkbookCommand.UngroupColumns("Budget", new ExcelColumnSpan(4, 6)),
        new WorkbookCommand.SetSheetPane("Budget", new ExcelSheetPane.Frozen(1, 2, 1, 2)),
        new WorkbookCommand.SetSheetZoom("Budget", 135),
        new WorkbookCommand.SetPrintLayout("Budget", defaultPrintLayout()),
        new WorkbookCommand.ClearPrintLayout("Budget"),
        new WorkbookCommand.SetCell("Budget", "A1", ExcelCellValue.date(LocalDate.of(2026, 3, 23))),
        new WorkbookCommand.SetRange(
            "Budget",
            "A1:B2",
            List.of(
                List.of(ExcelCellValue.text("Item"), ExcelCellValue.number(49.0d)),
                List.of(ExcelCellValue.text("Tax"), ExcelCellValue.number(10.0d)))),
        new WorkbookCommand.ClearRange("Budget", "C1:C2"),
        new WorkbookCommand.SetHyperlink(
            "Budget", "A1", new ExcelHyperlink.Url("https://example.com/report")),
        new WorkbookCommand.ClearHyperlink("Budget", "A1"),
        new WorkbookCommand.SetComment(
            "Budget", "A1", new ExcelComment("Review", "GridGrind", false)),
        new WorkbookCommand.ClearComment("Budget", "A1"),
        new WorkbookCommand.SetPicture(
            "Budget",
            new ExcelPictureDefinition(
                "BudgetPicture",
                new ExcelBinaryData(PNG_PIXEL_BYTES),
                ExcelPictureFormat.PNG,
                firstAnchor,
                "Queue preview")),
        new WorkbookCommand.SetChart("Budget", chartDefinition(firstAnchor)),
        new WorkbookCommand.SetPivotTable(
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
        new WorkbookCommand.SetShape(
            "Budget",
            new ExcelShapeDefinition(
                "BudgetShape",
                ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                firstAnchor,
                "rect",
                "Queue")),
        new WorkbookCommand.SetEmbeddedObject(
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
        new WorkbookCommand.SetDrawingObjectAnchor("Budget", "BudgetShape", movedAnchor),
        new WorkbookCommand.DeleteDrawingObject("Budget", "BudgetPicture"),
        new WorkbookCommand.ApplyStyle("Budget", "A1:B1", defaultStyle()),
        new WorkbookCommand.SetDataValidation("Budget", "B2:B5", validationDefinition()),
        new WorkbookCommand.ClearDataValidations(
            "Budget", new ExcelRangeSelection.Selected(List.of("C2:D4"))),
        new WorkbookCommand.SetConditionalFormatting(
            "Budget",
            new ExcelConditionalFormattingBlockDefinition(
                List.of("A2:A5"),
                List.of(
                    new ExcelConditionalFormattingRule.FormulaRule(
                        "A2>0",
                        true,
                        new ExcelDifferentialStyle(
                            "0.00", null, null, null, null, null, null, null, null))))),
        new WorkbookCommand.ClearConditionalFormatting(
            "Budget", new ExcelRangeSelection.Selected(List.of("A2:A5"))),
        new WorkbookCommand.SetAutofilter(
            "Budget",
            "A1:C4",
            List.of(
                new ExcelAutofilterFilterColumn(
                    0L, false, new ExcelAutofilterFilterCriterion.Values(List.of("Ada"), true))),
            new ExcelAutofilterSortState(
                "A1:C4",
                true,
                false,
                "",
                List.of(new ExcelAutofilterSortCondition("A2:A4", false, "", null, null)))),
        new WorkbookCommand.ClearAutofilter("Budget"),
        new WorkbookCommand.SetTable(
            new ExcelTableDefinition(
                "BudgetTable",
                "Budget",
                "A1:C4",
                true,
                new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false))),
        new WorkbookCommand.DeleteTable("BudgetTable", "Budget"),
        new WorkbookCommand.DeletePivotTable("Budget Pivot", "Budget"),
        new WorkbookCommand.SetNamedRange(
            new ExcelNamedRangeDefinition(
                "BudgetTotal",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Budget", "B4"))),
        new WorkbookCommand.DeleteNamedRange(
            "BudgetTotal", new ExcelNamedRangeScope.SheetScope("Budget")),
        new WorkbookCommand.AppendRow("Budget", List.of(ExcelCellValue.text("Item"))),
        new WorkbookCommand.AutoSizeColumns("Budget"));
  }

  static ExcelDrawingAnchor.TwoCell anchor(int fromRow, int fromColumn, int toRow, int toColumn) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromRow, fromColumn),
        new ExcelDrawingMarker(toRow, toColumn),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static ExcelChartDefinition chartDefinition(ExcelDrawingAnchor.TwoCell anchor) {
    return new ExcelChartDefinition.Bar(
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
                new ExcelChartDefinition.DataSource("A2:A4"),
                new ExcelChartDefinition.DataSource("B2:B4"))));
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
