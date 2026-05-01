package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookCommand sealed interface record construction. */
class WorkbookCommandTest {
  @Test
  void createsSupportedCommandsAndCopiesCollections() {
    List<ExcelCellValue> values = new ArrayList<>(List.of(ExcelCellValue.text("Item")));
    List<List<ExcelCellValue>> rows =
        new ArrayList<>(
            List.of(
                new ArrayList<>(List.of(ExcelCellValue.text("Item"), ExcelCellValue.number(49.0))),
                new ArrayList<>(List.of(ExcelCellValue.text("Tax"), ExcelCellValue.number(10.0)))));
    CreatedCommands commands = createSupportedCommands(values, rows, defaultStyle());

    values.clear();
    rows.clear();

    assertCreatedCommands(commands, defaultStyle());
  }

  @Test
  void buildsRowAndColumnStructureCommands() {
    WorkbookStructureCommand.InsertRows insertRows =
        new WorkbookStructureCommand.InsertRows("Budget", 2, 3);
    WorkbookStructureCommand.DeleteRows deleteRows =
        new WorkbookStructureCommand.DeleteRows("Budget", new ExcelRowSpan(4, 6));
    WorkbookStructureCommand.ShiftRows shiftRows =
        new WorkbookStructureCommand.ShiftRows("Budget", new ExcelRowSpan(1, 3), 2);
    WorkbookStructureCommand.InsertColumns insertColumns =
        new WorkbookStructureCommand.InsertColumns("Budget", 1, 2);
    WorkbookStructureCommand.DeleteColumns deleteColumns =
        new WorkbookStructureCommand.DeleteColumns("Budget", new ExcelColumnSpan(3, 4));
    WorkbookStructureCommand.ShiftColumns shiftColumns =
        new WorkbookStructureCommand.ShiftColumns("Budget", new ExcelColumnSpan(0, 1), -1);
    WorkbookStructureCommand.SetRowVisibility setRowVisibility =
        new WorkbookStructureCommand.SetRowVisibility("Budget", new ExcelRowSpan(5, 7), true);
    WorkbookStructureCommand.SetColumnVisibility setColumnVisibility =
        new WorkbookStructureCommand.SetColumnVisibility(
            "Budget", new ExcelColumnSpan(2, 3), false);
    WorkbookStructureCommand.GroupRows groupRows =
        new WorkbookStructureCommand.GroupRows("Budget", new ExcelRowSpan(8, 10), false);
    WorkbookStructureCommand.UngroupRows ungroupRows =
        new WorkbookStructureCommand.UngroupRows("Budget", new ExcelRowSpan(8, 10));
    WorkbookStructureCommand.GroupColumns groupColumns =
        new WorkbookStructureCommand.GroupColumns("Budget", new ExcelColumnSpan(4, 6), true);
    WorkbookStructureCommand.UngroupColumns ungroupColumns =
        new WorkbookStructureCommand.UngroupColumns("Budget", new ExcelColumnSpan(4, 6));

    assertEquals(2, insertRows.rowIndex());
    assertEquals(new ExcelRowSpan(4, 6), deleteRows.rows());
    assertEquals(2, shiftRows.delta());
    assertEquals(1, insertColumns.columnIndex());
    assertEquals(new ExcelColumnSpan(3, 4), deleteColumns.columns());
    assertEquals(-1, shiftColumns.delta());
    assertTrue(setRowVisibility.hidden());
    assertFalse(setColumnVisibility.hidden());
    assertFalse(groupRows.collapsed());
    assertEquals(new ExcelRowSpan(8, 10), ungroupRows.rows());
    assertTrue(groupColumns.collapsed());
    assertEquals(new ExcelColumnSpan(4, 6), ungroupColumns.columns());
  }

  @Test
  void buildsWorkbookProtectionAndAdvancedAutofilterCommands() {
    List<ExcelAutofilterFilterColumn> criteria =
        new ArrayList<>(
            List.of(
                new ExcelAutofilterFilterColumn(
                    0L, false, new ExcelAutofilterFilterCriterion.Values(List.of("Ada"), true))));
    WorkbookSheetCommand.SetSheetProtection sheetProtection =
        new WorkbookSheetCommand.SetSheetProtection("Budget", protectionSettings(), "gridgrind");
    WorkbookSheetCommand.SetWorkbookProtection workbookProtection =
        new WorkbookSheetCommand.SetWorkbookProtection(
            new ExcelWorkbookProtectionSettings(true, false, true, "book", "review"));
    WorkbookSheetCommand.ClearWorkbookProtection clearWorkbookProtection =
        new WorkbookSheetCommand.ClearWorkbookProtection();
    WorkbookTabularCommand.SetAutofilter setAutofilter =
        new WorkbookTabularCommand.SetAutofilter(
            "Budget",
            "A1:C4",
            criteria,
            new ExcelAutofilterSortState(
                "A1:C4",
                true,
                false,
                java.util.Optional.empty(),
                List.of(new ExcelAutofilterSortCondition.Value("A2:A4", false))));

    criteria.clear();

    assertEquals("gridgrind", sheetProtection.password());
    assertEquals("book", workbookProtection.protection().workbookPassword());
    assertNotNull(clearWorkbookProtection);
    assertEquals(1, setAutofilter.criteria().size());
    assertFalse(setAutofilter.criteria().getFirst().showButton());
    assertTrue(setAutofilter.sortState().caseSensitive());

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetCommand.SetSheetProtection("Budget", protectionSettings(), " "));
  }

  @Test
  void deletePivotTableCommandRejectsInvalidInputs() {
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookTabularCommand.DeletePivotTable(null, "Budget"));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookTabularCommand.DeletePivotTable("Budget Pivot", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookTabularCommand.DeletePivotTable("Budget Pivot", " "));
  }

  private CreatedCommands createSupportedCommands(
      List<ExcelCellValue> values, List<List<ExcelCellValue>> rows, ExcelCellStyle style) {
    return new CreatedCommands(
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
        new WorkbookStructureCommand.MergeCells("Budget", "A1:B2"),
        new WorkbookStructureCommand.UnmergeCells("Budget", "A1:B2"),
        new WorkbookStructureCommand.SetColumnWidth("Budget", 0, 1, 16.0),
        new WorkbookStructureCommand.SetRowHeight("Budget", 0, 2, 28.5),
        new WorkbookLayoutCommand.SetSheetPane("Budget", new ExcelSheetPane.Frozen(1, 2, 1, 2)),
        new WorkbookLayoutCommand.SetSheetZoom("Budget", 135),
        new WorkbookLayoutCommand.SetPrintLayout("Budget", defaultPrintLayout()),
        new WorkbookLayoutCommand.ClearPrintLayout("Budget"),
        new WorkbookCellCommand.SetCell(
            "Budget", "A1", ExcelCellValue.date(LocalDate.of(2026, 3, 23))),
        new WorkbookCellCommand.SetRange("Budget", "A1:B2", rows),
        new WorkbookCellCommand.ClearRange("Budget", "C1:C2"),
        new WorkbookAnnotationCommand.SetHyperlink(
            "Budget", "A1", new ExcelHyperlink.Url("https://example.com/report")),
        new WorkbookAnnotationCommand.ClearHyperlink("Budget", "A1"),
        new WorkbookAnnotationCommand.SetComment(
            "Budget", "A1", new ExcelComment("Review", "GridGrind", false)),
        new WorkbookAnnotationCommand.ClearComment("Budget", "A1"),
        new WorkbookFormattingCommand.ApplyStyle("Budget", "A1:B1", style),
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
        new WorkbookTabularCommand.SetAutofilter("Budget", "A1:C4"),
        new WorkbookTabularCommand.ClearAutofilter("Budget"),
        new WorkbookTabularCommand.SetTable(
            new ExcelTableDefinition(
                "BudgetTable",
                "Budget",
                "A1:C4",
                true,
                new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false))),
        new WorkbookTabularCommand.DeleteTable("BudgetTable", "Budget"),
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
        new WorkbookTabularCommand.DeletePivotTable("Budget Pivot", "Budget"),
        new WorkbookMetadataCommand.SetNamedRange(
            new ExcelNamedRangeDefinition(
                "BudgetTotal",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Budget", "B4"))),
        new WorkbookMetadataCommand.DeleteNamedRange(
            "BudgetTotal", new ExcelNamedRangeScope.SheetScope("Budget")),
        new WorkbookCellCommand.AppendRow("Budget", values),
        new WorkbookLayoutCommand.AutoSizeColumns("Budget"));
  }

  private void assertCreatedCommands(CreatedCommands commands, ExcelCellStyle style) {
    assertEquals("Budget", commands.createSheet().sheetName());
    assertEquals("Summary", commands.renameSheet().newSheetName());
    assertEquals("Archive", commands.deleteSheet().sheetName());
    assertEquals(1, commands.moveSheet().targetIndex());
    assertEquals("Budget Copy", commands.copySheet().newSheetName());
    assertEquals("Budget", commands.setActiveSheet().sheetName());
    assertEquals(List.of("Budget", "Summary"), commands.setSelectedSheets().sheetNames());
    assertEquals(ExcelSheetVisibility.VERY_HIDDEN, commands.setSheetVisibility().visibility());
    assertEquals(protectionSettings(), commands.setSheetProtection().protection());
    assertEquals("Budget", commands.clearSheetProtection().sheetName());
    assertEquals("A1:B2", commands.mergeCells().range());
    assertEquals("A1:B2", commands.unmergeCells().range());
    assertEquals(16.0, commands.setColumnWidth().widthCharacters());
    assertEquals(28.5, commands.setRowHeight().heightPoints());
    assertEquals(new ExcelSheetPane.Frozen(1, 2, 1, 2), commands.setSheetPane().pane());
    assertEquals(135, commands.setSheetZoom().zoomPercent());
    assertEquals(defaultPrintLayout(), commands.setPrintLayout().printLayout());
    assertEquals("Budget", commands.clearPrintLayout().sheetName());
    assertEquals("A1", commands.setCell().address());
    assertEquals("A1:B2", commands.setRange().range());
    assertEquals(2, commands.setRange().rows().size());
    assertEquals("C1:C2", commands.clearRange().range());
    assertEquals(ExcelHyperlinkType.URL, commands.setHyperlink().target().type());
    assertEquals("A1", commands.clearHyperlink().address());
    assertEquals("Review", commands.setComment().comment().text());
    assertEquals("A1", commands.clearComment().address());
    assertEquals(style, commands.applyStyle().style());
    assertEquals("B2:B5", commands.setDataValidation().range());
    assertEquals(
        List.of("C2:D4"),
        ((ExcelRangeSelection.Selected) commands.clearDataValidations().selection()).ranges());
    assertEquals("Budget", commands.setConditionalFormatting().sheetName());
    assertEquals(List.of("A2:A5"), commands.setConditionalFormatting().block().ranges());
    assertEquals(
        List.of("A2:A5"),
        ((ExcelRangeSelection.Selected) commands.clearConditionalFormatting().selection())
            .ranges());
    assertEquals("A1:C4", commands.setAutofilter().range());
    assertEquals("Budget", commands.clearAutofilter().sheetName());
    assertEquals("BudgetTable", commands.setTable().definition().name());
    assertEquals("Budget", commands.deleteTable().sheetName());
    assertEquals("Budget Pivot", commands.setPivotTable().definition().name());
    assertEquals("Budget", commands.deletePivotTable().sheetName());
    assertEquals("BudgetTotal", commands.setNamedRange().definition().name());
    assertEquals(
        "Budget",
        ((ExcelNamedRangeScope.SheetScope) commands.deleteNamedRange().scope()).sheetName());
    assertEquals(1, commands.appendRow().values().size());
    assertEquals("Budget", commands.autoSizeColumns().sheetName());
  }

  @Test
  void validatesFormulaCellTargetInputs() {
    assertThrows(NullPointerException.class, () -> new ExcelFormulaCellTarget(null, "A1"));
    assertThrows(IllegalArgumentException.class, () -> new ExcelFormulaCellTarget(" ", "A1"));
    assertThrows(NullPointerException.class, () -> new ExcelFormulaCellTarget("Budget", null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelFormulaCellTarget("Budget", " "));
    assertEquals(
        new ExcelFormulaCellTarget("Budget", "C2"), new ExcelFormulaCellTarget("Budget", "C2"));
  }

  @Test
  void validatesSheetIdentityCommandInputs() {
    assertThrows(NullPointerException.class, () -> new WorkbookSheetCommand.CreateSheet(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookSheetCommand.CreateSheet(" "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookSheetCommand.RenameSheet(null, "New"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookSheetCommand.RenameSheet(" ", "New"));
    assertThrows(
        NullPointerException.class, () -> new WorkbookSheetCommand.RenameSheet("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookSheetCommand.RenameSheet("Budget", " "));
    assertThrows(NullPointerException.class, () -> new WorkbookSheetCommand.DeleteSheet(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookSheetCommand.DeleteSheet(" "));
    assertThrows(NullPointerException.class, () -> new WorkbookSheetCommand.MoveSheet(null, 0));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookSheetCommand.MoveSheet(" ", 0));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookSheetCommand.MoveSheet("Budget", -1));
  }

  @Test
  void validatesSheetCopyAndStateCommandInputs() {
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookSheetCommand.CopySheet(
                null, "Budget Copy", new ExcelSheetCopyPosition.AppendAtEnd()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSheetCommand.CopySheet(
                " ", "Budget Copy", new ExcelSheetCopyPosition.AppendAtEnd()));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookSheetCommand.CopySheet(
                "Budget", null, new ExcelSheetCopyPosition.AppendAtEnd()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSheetCommand.CopySheet(
                "Budget", " ", new ExcelSheetCopyPosition.AppendAtEnd()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookSheetCommand.CopySheet("Budget", "Budget Copy", null));
    assertThrows(NullPointerException.class, () -> new WorkbookSheetCommand.SetActiveSheet(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookSheetCommand.SetActiveSheet(" "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookSheetCommand.SetSelectedSheets(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetCommand.SetSelectedSheets(List.of()));
    List<String> selectedSheetsWithNull = new ArrayList<>();
    selectedSheetsWithNull.add(null);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookSheetCommand.SetSelectedSheets(selectedSheetsWithNull));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetCommand.SetSelectedSheets(List.of("Budget", " ")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetCommand.SetSelectedSheets(List.of("Budget", "Budget")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetCommand.SetSheetVisibility(" ", ExcelSheetVisibility.HIDDEN));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookSheetCommand.SetSheetVisibility("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetCommand.SetSheetProtection(" ", protectionSettings()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookSheetCommand.SetSheetProtection("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookSheetCommand.ClearSheetProtection(" "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookSheetCommand.ClearSheetProtection(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.InsertRows("Budget", -1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.InsertRows("Budget", ExcelRowSpan.MAX_ROW_INDEX + 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.ShiftRows("Budget", new ExcelRowSpan(0, 0), 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.InsertColumns("Budget", -1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookStructureCommand.InsertColumns(
                "Budget", ExcelColumnSpan.MAX_COLUMN_INDEX + 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.ShiftColumns("Budget", new ExcelColumnSpan(0, 0), 0));
  }

  @Test
  void validatesStructuralCommandSheetNamesAndCounts() {
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookStructureCommand.InsertRows(" ", 0, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.InsertRows("Budget", 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.DeleteRows(" ", new ExcelRowSpan(0, 0)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.ShiftRows(" ", new ExcelRowSpan(0, 0), 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.SetRowVisibility(" ", new ExcelRowSpan(0, 0), true));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.GroupRows(" ", new ExcelRowSpan(0, 0), false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.UngroupRows(" ", new ExcelRowSpan(0, 0)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.InsertColumns(" ", 0, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.InsertColumns("Budget", 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.DeleteColumns(" ", new ExcelColumnSpan(0, 0)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.ShiftColumns(" ", new ExcelColumnSpan(0, 0), 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookStructureCommand.SetColumnVisibility(" ", new ExcelColumnSpan(0, 0), true));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.GroupColumns(" ", new ExcelColumnSpan(0, 0), false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.UngroupColumns(" ", new ExcelColumnSpan(0, 0)));
  }

  @Test
  void validatesMergeAndLayoutSizingCommandInputs() {
    assertThrows(
        NullPointerException.class, () -> new WorkbookStructureCommand.MergeCells(null, "A1:B2"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookStructureCommand.MergeCells(" ", "A1"));
    assertThrows(
        NullPointerException.class, () -> new WorkbookStructureCommand.MergeCells("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.MergeCells("Budget", " "));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookStructureCommand.UnmergeCells("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.UnmergeCells(" ", "A1:B2"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.UnmergeCells("Budget", " "));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookStructureCommand.SetColumnWidth(null, 0, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.SetColumnWidth(" ", 0, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.SetColumnWidth("Budget", -1, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.SetColumnWidth("Budget", 0, -1, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.SetColumnWidth("Budget", 2, 1, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.SetColumnWidth("Budget", 0, 0, 256.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.SetColumnWidth("Budget", 0, 0, Double.MIN_VALUE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookStructureCommand.SetColumnWidth("Budget", 0, 0, Double.POSITIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.SetRowHeight("Budget", -1, 0, 28.5));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookStructureCommand.SetRowHeight(null, 0, 0, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.SetRowHeight(" ", 0, 0, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.SetRowHeight("Budget", 0, -1, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.SetRowHeight("Budget", 2, 1, 28.5));
    assertDoesNotThrow(
        () ->
            new WorkbookStructureCommand.SetRowHeight(
                "Budget", 0, 0, ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookStructureCommand.SetRowHeight(
                "Budget", 0, 0, Math.nextUp(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookStructureCommand.SetRowHeight(
                "Budget", 0, 0, ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS + 1.0d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.SetRowHeight("Budget", 0, 0, Double.MIN_VALUE));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookStructureCommand.SetRowHeight("Budget", 0, 0, 0.0));
  }

  @Test
  void validatesPaneZoomAndPrintLayoutCommandInputs() {
    assertDoesNotThrow(
        () -> new WorkbookLayoutCommand.SetSheetPane("Budget", new ExcelSheetPane.None()));
    assertDoesNotThrow(
        () ->
            new WorkbookLayoutCommand.SetSheetPane(
                "Budget", new ExcelSheetPane.Frozen(0, 2, 0, 2)));
    assertDoesNotThrow(
        () ->
            new WorkbookLayoutCommand.SetSheetPane(
                "Budget", new ExcelSheetPane.Split(1200, 0, 3, 0, ExcelPaneRegion.UPPER_RIGHT)));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookLayoutCommand.SetSheetPane(null, new ExcelSheetPane.None()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookLayoutCommand.SetSheetPane(" ", new ExcelSheetPane.None()));
    assertThrows(
        NullPointerException.class, () -> new WorkbookLayoutCommand.SetSheetPane("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookLayoutCommand.SetSheetZoom("Budget", 9));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookLayoutCommand.SetSheetZoom("Budget", 401));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookLayoutCommand.SetSheetZoom(" ", 100));
    assertThrows(
        NullPointerException.class, () -> new WorkbookLayoutCommand.SetPrintLayout("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookLayoutCommand.SetPrintLayout(" ", defaultPrintLayout()));
    WorkbookLayoutCommand.SetSheetPresentation setSheetPresentation =
        assertDoesNotThrow(
            () ->
                new WorkbookLayoutCommand.SetSheetPresentation(
                    "Budget", ExcelSheetPresentation.defaults()));
    assertEquals("Budget", setSheetPresentation.sheetName());
    assertEquals(ExcelSheetPresentation.defaults(), setSheetPresentation.presentation());
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookLayoutCommand.SetSheetPresentation(
                null, ExcelSheetPresentation.defaults()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookLayoutCommand.SetSheetPresentation(" ", ExcelSheetPresentation.defaults()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookLayoutCommand.SetSheetPresentation("Budget", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookLayoutCommand.ClearPrintLayout(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookLayoutCommand.ClearPrintLayout(" "));
  }

  @Test
  void validatesCellRangeAndLinkCommentCommandInputs() {
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCellCommand.SetCell(null, "A1", ExcelCellValue.text("x")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCellCommand.SetCell("Budget", null, ExcelCellValue.text("x")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCellCommand.SetCell("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCellCommand.SetCell(" ", "A1", ExcelCellValue.text("x")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCellCommand.SetCell("Budget", " ", ExcelCellValue.text("x")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCellCommand.SetRange(null, "A1", List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCellCommand.SetRange(" ", "A1", List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCellCommand.SetRange("Budget", null, List.of()));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCellCommand.SetRange("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCellCommand.SetRange("Budget", " ", List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCellCommand.SetRange("Budget", "A1", List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCellCommand.SetRange("Budget", "A1:B2", List.of(List.of())));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCellCommand.SetRange(
                "Budget",
                "A1:B2",
                List.of(
                    List.of(ExcelCellValue.text("x")),
                    List.of(ExcelCellValue.text("y"), ExcelCellValue.text("z")))));
    List<List<ExcelCellValue>> rowsWithNull = new ArrayList<>();
    List<ExcelCellValue> rowWithNull = new ArrayList<>();
    rowWithNull.add(null);
    rowsWithNull.add(rowWithNull);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCellCommand.SetRange("Budget", "A1", rowsWithNull));
    assertThrows(NullPointerException.class, () -> new WorkbookCellCommand.ClearRange(null, "A1"));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCellCommand.ClearRange("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCellCommand.ClearRange(" ", "A1"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCellCommand.ClearRange("Budget", " "));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookAnnotationCommand.SetHyperlink(
                null, "A1", new ExcelHyperlink.Url("https://example.com")));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookAnnotationCommand.SetHyperlink(
                "Budget", null, new ExcelHyperlink.Url("https://example.com")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnnotationCommand.SetHyperlink("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookAnnotationCommand.SetHyperlink(
                " ", "A1", new ExcelHyperlink.Url("https://example.com")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookAnnotationCommand.SetHyperlink(
                "Budget", " ", new ExcelHyperlink.Url("https://example.com")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookAnnotationCommand.ClearHyperlink(null, "A1"));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnnotationCommand.ClearHyperlink("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookAnnotationCommand.ClearHyperlink(" ", "A1"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookAnnotationCommand.ClearHyperlink("Budget", " "));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookAnnotationCommand.SetComment(
                null, "A1", new ExcelComment("Review", "GridGrind", false)));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookAnnotationCommand.SetComment(
                "Budget", null, new ExcelComment("Review", "GridGrind", false)));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnnotationCommand.SetComment("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookAnnotationCommand.SetComment(
                " ", "A1", new ExcelComment("Review", "GridGrind", false)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookAnnotationCommand.SetComment(
                "Budget", " ", new ExcelComment("Review", "GridGrind", false)));
    assertThrows(
        NullPointerException.class, () -> new WorkbookAnnotationCommand.ClearComment(null, "A1"));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnnotationCommand.ClearComment("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookAnnotationCommand.ClearComment(" ", "A1"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookAnnotationCommand.ClearComment("Budget", " "));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookFormattingCommand.ApplyStyle(null, "A1", ExcelCellStyle.numberFormat("0")));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookFormattingCommand.ApplyStyle(
                "Budget", null, ExcelCellStyle.numberFormat("0")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookFormattingCommand.ApplyStyle("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookFormattingCommand.ApplyStyle(" ", "A1", ExcelCellStyle.numberFormat("0")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookFormattingCommand.ApplyStyle(
                "Budget", " ", ExcelCellStyle.numberFormat("0")));
  }

  @Test
  void validatesValidationNamedRangeAndAppendCommandInputs() {
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookFormattingCommand.SetDataValidation("Budget", "A1", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookFormattingCommand.SetDataValidation(null, "A1", validationDefinition()));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookFormattingCommand.SetDataValidation(
                "Budget", null, validationDefinition()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookFormattingCommand.SetDataValidation(" ", "A1", validationDefinition()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookFormattingCommand.SetDataValidation("Budget", " ", validationDefinition()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookFormattingCommand.ClearDataValidations("Budget", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookFormattingCommand.ClearDataValidations(
                null, new ExcelRangeSelection.All()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookFormattingCommand.ClearDataValidations(" ", new ExcelRangeSelection.All()));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookFormattingCommand.SetConditionalFormatting(
                null,
                new ExcelConditionalFormattingBlockDefinition(
                    List.of("A2:A5"),
                    List.of(
                        new ExcelConditionalFormattingRule.FormulaRule(
                            "A2>0",
                            true,
                            new ExcelDifferentialStyle(
                                "0.00", null, null, null, null, null, null, null, null))))));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookFormattingCommand.SetConditionalFormatting("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookFormattingCommand.SetConditionalFormatting(
                " ",
                new ExcelConditionalFormattingBlockDefinition(
                    List.of("A2:A5"),
                    List.of(
                        new ExcelConditionalFormattingRule.FormulaRule(
                            "A2>0",
                            true,
                            new ExcelDifferentialStyle(
                                "0.00", null, null, null, null, null, null, null, null))))));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookFormattingCommand.ClearConditionalFormatting("Budget", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookFormattingCommand.ClearConditionalFormatting(
                null, new ExcelRangeSelection.All()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookFormattingCommand.ClearConditionalFormatting(
                " ", new ExcelRangeSelection.All()));
    assertThrows(
        NullPointerException.class, () -> new WorkbookTabularCommand.SetAutofilter(null, "A1:B2"));
    assertThrows(
        NullPointerException.class, () -> new WorkbookTabularCommand.SetAutofilter("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookTabularCommand.SetAutofilter(" ", "A1:B2"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookTabularCommand.SetAutofilter("Budget", " "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookTabularCommand.ClearAutofilter(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookTabularCommand.ClearAutofilter(" "));
    assertThrows(NullPointerException.class, () -> new WorkbookTabularCommand.SetTable(null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookTabularCommand.DeleteTable(null, "Budget"));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookTabularCommand.DeleteTable("BudgetTable", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookTabularCommand.DeleteTable(" ", "Budget"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookTabularCommand.DeleteTable("BudgetTable", " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookTabularCommand.SetTable(
                new ExcelTableDefinition(
                    "A1", "Budget", "A1:B2", false, new ExcelTableStyle.None())));
    assertThrows(NullPointerException.class, () -> new WorkbookMetadataCommand.SetNamedRange(null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookMetadataCommand.DeleteNamedRange(
                null, new ExcelNamedRangeScope.WorkbookScope()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookMetadataCommand.DeleteNamedRange("BudgetTotal", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookMetadataCommand.DeleteNamedRange(
                "_XLNM.PRINT_AREA", new ExcelNamedRangeScope.WorkbookScope()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookTabularCommand.DeleteTable("A1", "Budget"));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCellCommand.AppendRow(null, List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCellCommand.AppendRow(" ", List.of()));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCellCommand.AppendRow("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCellCommand.AppendRow("Budget", List.of()));
    List<ExcelCellValue> valuesWithNull = new ArrayList<>();
    valuesWithNull.add(null);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCellCommand.AppendRow("Budget", valuesWithNull));
    assertThrows(NullPointerException.class, () -> new WorkbookLayoutCommand.AutoSizeColumns(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookLayoutCommand.AutoSizeColumns(" "));
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

  private ExcelCellStyle defaultStyle() {
    return new ExcelCellStyle(
        "#,##0.00",
        new ExcelCellAlignment(
            true, ExcelHorizontalAlignment.RIGHT, ExcelVerticalAlignment.CENTER, null, null),
        new ExcelCellFont(true, false, null, null, null, null, null),
        null,
        null,
        null);
  }

  private record CreatedCommands(
      WorkbookSheetCommand.CreateSheet createSheet,
      WorkbookSheetCommand.RenameSheet renameSheet,
      WorkbookSheetCommand.DeleteSheet deleteSheet,
      WorkbookSheetCommand.MoveSheet moveSheet,
      WorkbookSheetCommand.CopySheet copySheet,
      WorkbookSheetCommand.SetActiveSheet setActiveSheet,
      WorkbookSheetCommand.SetSelectedSheets setSelectedSheets,
      WorkbookSheetCommand.SetSheetVisibility setSheetVisibility,
      WorkbookSheetCommand.SetSheetProtection setSheetProtection,
      WorkbookSheetCommand.ClearSheetProtection clearSheetProtection,
      WorkbookStructureCommand.MergeCells mergeCells,
      WorkbookStructureCommand.UnmergeCells unmergeCells,
      WorkbookStructureCommand.SetColumnWidth setColumnWidth,
      WorkbookStructureCommand.SetRowHeight setRowHeight,
      WorkbookLayoutCommand.SetSheetPane setSheetPane,
      WorkbookLayoutCommand.SetSheetZoom setSheetZoom,
      WorkbookLayoutCommand.SetPrintLayout setPrintLayout,
      WorkbookLayoutCommand.ClearPrintLayout clearPrintLayout,
      WorkbookCellCommand.SetCell setCell,
      WorkbookCellCommand.SetRange setRange,
      WorkbookCellCommand.ClearRange clearRange,
      WorkbookAnnotationCommand.SetHyperlink setHyperlink,
      WorkbookAnnotationCommand.ClearHyperlink clearHyperlink,
      WorkbookAnnotationCommand.SetComment setComment,
      WorkbookAnnotationCommand.ClearComment clearComment,
      WorkbookFormattingCommand.ApplyStyle applyStyle,
      WorkbookFormattingCommand.SetDataValidation setDataValidation,
      WorkbookFormattingCommand.ClearDataValidations clearDataValidations,
      WorkbookFormattingCommand.SetConditionalFormatting setConditionalFormatting,
      WorkbookFormattingCommand.ClearConditionalFormatting clearConditionalFormatting,
      WorkbookTabularCommand.SetAutofilter setAutofilter,
      WorkbookTabularCommand.ClearAutofilter clearAutofilter,
      WorkbookTabularCommand.SetTable setTable,
      WorkbookTabularCommand.DeleteTable deleteTable,
      WorkbookTabularCommand.SetPivotTable setPivotTable,
      WorkbookTabularCommand.DeletePivotTable deletePivotTable,
      WorkbookMetadataCommand.SetNamedRange setNamedRange,
      WorkbookMetadataCommand.DeleteNamedRange deleteNamedRange,
      WorkbookCellCommand.AppendRow appendRow,
      WorkbookLayoutCommand.AutoSizeColumns autoSizeColumns) {}

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
