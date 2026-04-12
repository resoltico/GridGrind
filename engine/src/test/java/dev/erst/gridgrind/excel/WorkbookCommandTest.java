package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

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
    WorkbookCommand.InsertRows insertRows = new WorkbookCommand.InsertRows("Budget", 2, 3);
    WorkbookCommand.DeleteRows deleteRows =
        new WorkbookCommand.DeleteRows("Budget", new ExcelRowSpan(4, 6));
    WorkbookCommand.ShiftRows shiftRows =
        new WorkbookCommand.ShiftRows("Budget", new ExcelRowSpan(1, 3), 2);
    WorkbookCommand.InsertColumns insertColumns = new WorkbookCommand.InsertColumns("Budget", 1, 2);
    WorkbookCommand.DeleteColumns deleteColumns =
        new WorkbookCommand.DeleteColumns("Budget", new ExcelColumnSpan(3, 4));
    WorkbookCommand.ShiftColumns shiftColumns =
        new WorkbookCommand.ShiftColumns("Budget", new ExcelColumnSpan(0, 1), -1);
    WorkbookCommand.SetRowVisibility setRowVisibility =
        new WorkbookCommand.SetRowVisibility("Budget", new ExcelRowSpan(5, 7), true);
    WorkbookCommand.SetColumnVisibility setColumnVisibility =
        new WorkbookCommand.SetColumnVisibility("Budget", new ExcelColumnSpan(2, 3), false);
    WorkbookCommand.GroupRows groupRows =
        new WorkbookCommand.GroupRows("Budget", new ExcelRowSpan(8, 10), false);
    WorkbookCommand.UngroupRows ungroupRows =
        new WorkbookCommand.UngroupRows("Budget", new ExcelRowSpan(8, 10));
    WorkbookCommand.GroupColumns groupColumns =
        new WorkbookCommand.GroupColumns("Budget", new ExcelColumnSpan(4, 6), true);
    WorkbookCommand.UngroupColumns ungroupColumns =
        new WorkbookCommand.UngroupColumns("Budget", new ExcelColumnSpan(4, 6));

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
    WorkbookCommand.SetSheetProtection sheetProtection =
        new WorkbookCommand.SetSheetProtection("Budget", protectionSettings(), "gridgrind");
    WorkbookCommand.SetWorkbookProtection workbookProtection =
        new WorkbookCommand.SetWorkbookProtection(
            new ExcelWorkbookProtectionSettings(true, false, true, "book", "review"));
    WorkbookCommand.ClearWorkbookProtection clearWorkbookProtection =
        new WorkbookCommand.ClearWorkbookProtection();
    WorkbookCommand.SetAutofilter setAutofilter =
        new WorkbookCommand.SetAutofilter(
            "Budget",
            "A1:C4",
            criteria,
            new ExcelAutofilterSortState(
                "A1:C4",
                true,
                false,
                "",
                List.of(new ExcelAutofilterSortCondition("A2:A4", false, "", null, null))));

    criteria.clear();

    assertEquals("gridgrind", sheetProtection.password());
    assertEquals("book", workbookProtection.protection().workbookPassword());
    assertNotNull(clearWorkbookProtection);
    assertEquals(1, setAutofilter.criteria().size());
    assertFalse(setAutofilter.criteria().getFirst().showButton());
    assertTrue(setAutofilter.sortState().caseSensitive());

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetSheetProtection("Budget", protectionSettings(), " "));
  }

  @Test
  void deletePivotTableCommandRejectsInvalidInputs() {
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.DeletePivotTable(null, "Budget"));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.DeletePivotTable("Budget Pivot", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.DeletePivotTable("Budget Pivot", " "));
  }

  private CreatedCommands createSupportedCommands(
      List<ExcelCellValue> values, List<List<ExcelCellValue>> rows, ExcelCellStyle style) {
    return new CreatedCommands(
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
        new WorkbookCommand.MergeCells("Budget", "A1:B2"),
        new WorkbookCommand.UnmergeCells("Budget", "A1:B2"),
        new WorkbookCommand.SetColumnWidth("Budget", 0, 1, 16.0),
        new WorkbookCommand.SetRowHeight("Budget", 0, 2, 28.5),
        new WorkbookCommand.SetSheetPane("Budget", new ExcelSheetPane.Frozen(1, 2, 1, 2)),
        new WorkbookCommand.SetSheetZoom("Budget", 135),
        new WorkbookCommand.SetPrintLayout("Budget", defaultPrintLayout()),
        new WorkbookCommand.ClearPrintLayout("Budget"),
        new WorkbookCommand.SetCell("Budget", "A1", ExcelCellValue.date(LocalDate.of(2026, 3, 23))),
        new WorkbookCommand.SetRange("Budget", "A1:B2", rows),
        new WorkbookCommand.ClearRange("Budget", "C1:C2"),
        new WorkbookCommand.SetHyperlink(
            "Budget", "A1", new ExcelHyperlink.Url("https://example.com/report")),
        new WorkbookCommand.ClearHyperlink("Budget", "A1"),
        new WorkbookCommand.SetComment(
            "Budget", "A1", new ExcelComment("Review", "GridGrind", false)),
        new WorkbookCommand.ClearComment("Budget", "A1"),
        new WorkbookCommand.ApplyStyle("Budget", "A1:B1", style),
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
        new WorkbookCommand.SetAutofilter("Budget", "A1:C4"),
        new WorkbookCommand.ClearAutofilter("Budget"),
        new WorkbookCommand.SetTable(
            new ExcelTableDefinition(
                "BudgetTable",
                "Budget",
                "A1:C4",
                true,
                new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false))),
        new WorkbookCommand.DeleteTable("BudgetTable", "Budget"),
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
        new WorkbookCommand.DeletePivotTable("Budget Pivot", "Budget"),
        new WorkbookCommand.SetNamedRange(
            new ExcelNamedRangeDefinition(
                "BudgetTotal",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Budget", "B4"))),
        new WorkbookCommand.DeleteNamedRange(
            "BudgetTotal", new ExcelNamedRangeScope.SheetScope("Budget")),
        new WorkbookCommand.AppendRow("Budget", values),
        new WorkbookCommand.AutoSizeColumns("Budget"),
        new WorkbookCommand.EvaluateAllFormulas(),
        new WorkbookCommand.EvaluateFormulaCells(
            List.of(new ExcelFormulaCellTarget("Budget", "B4"))),
        new WorkbookCommand.ClearFormulaCaches(),
        new WorkbookCommand.ForceFormulaRecalculationOnOpen());
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
    assertNotNull(commands.evaluate());
    assertEquals(
        List.of(new ExcelFormulaCellTarget("Budget", "B4")),
        commands.evaluateFormulaCells().cells());
    assertNotNull(commands.clearFormulaCaches());
    assertNotNull(commands.recalc());
  }

  @Test
  void validatesFormulaLifecycleCommandInputs() {
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.EvaluateFormulaCells(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.EvaluateFormulaCells(List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.EvaluateFormulaCells(List.of((ExcelFormulaCellTarget) null)));
    assertThrows(NullPointerException.class, () -> new ExcelFormulaCellTarget(null, "A1"));
    assertThrows(IllegalArgumentException.class, () -> new ExcelFormulaCellTarget(" ", "A1"));
    assertThrows(NullPointerException.class, () -> new ExcelFormulaCellTarget("Budget", null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelFormulaCellTarget("Budget", " "));

    WorkbookCommand.EvaluateFormulaCells evaluateFormulaCells =
        new WorkbookCommand.EvaluateFormulaCells(
            List.of(new ExcelFormulaCellTarget("Budget", "C2")));
    WorkbookCommand.ClearFormulaCaches clearFormulaCaches =
        new WorkbookCommand.ClearFormulaCaches();

    assertEquals(List.of(new ExcelFormulaCellTarget("Budget", "C2")), evaluateFormulaCells.cells());
    assertNotNull(clearFormulaCaches);
  }

  @Test
  void validatesSheetIdentityCommandInputs() {
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.CreateSheet(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.CreateSheet(" "));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.RenameSheet(null, "New"));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.RenameSheet(" ", "New"));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.RenameSheet("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.RenameSheet("Budget", " "));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.DeleteSheet(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.DeleteSheet(" "));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.MoveSheet(null, 0));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.MoveSheet(" ", 0));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.MoveSheet("Budget", -1));
  }

  @Test
  void validatesSheetCopyAndStateCommandInputs() {
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookCommand.CopySheet(
                null, "Budget Copy", new ExcelSheetCopyPosition.AppendAtEnd()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCommand.CopySheet(
                " ", "Budget Copy", new ExcelSheetCopyPosition.AppendAtEnd()));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookCommand.CopySheet(
                "Budget", null, new ExcelSheetCopyPosition.AppendAtEnd()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCommand.CopySheet("Budget", " ", new ExcelSheetCopyPosition.AppendAtEnd()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.CopySheet("Budget", "Budget Copy", null));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.SetActiveSheet(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.SetActiveSheet(" "));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.SetSelectedSheets(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetSelectedSheets(List.of()));
    List<String> selectedSheetsWithNull = new ArrayList<>();
    selectedSheetsWithNull.add(null);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetSelectedSheets(selectedSheetsWithNull));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetSelectedSheets(List.of("Budget", " ")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetSelectedSheets(List.of("Budget", "Budget")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetSheetVisibility(" ", ExcelSheetVisibility.HIDDEN));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetSheetVisibility("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetSheetProtection(" ", protectionSettings()));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetSheetProtection("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.ClearSheetProtection(" "));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.ClearSheetProtection(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.InsertRows("Budget", -1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.InsertRows("Budget", ExcelRowSpan.MAX_ROW_INDEX + 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.ShiftRows("Budget", new ExcelRowSpan(0, 0), 0));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.InsertColumns("Budget", -1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.InsertColumns("Budget", ExcelColumnSpan.MAX_COLUMN_INDEX + 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.ShiftColumns("Budget", new ExcelColumnSpan(0, 0), 0));
  }

  @Test
  void validatesStructuralCommandSheetNamesAndCounts() {
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.InsertRows(" ", 0, 1));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.InsertRows("Budget", 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.DeleteRows(" ", new ExcelRowSpan(0, 0)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.ShiftRows(" ", new ExcelRowSpan(0, 0), 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRowVisibility(" ", new ExcelRowSpan(0, 0), true));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.GroupRows(" ", new ExcelRowSpan(0, 0), false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.UngroupRows(" ", new ExcelRowSpan(0, 0)));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.InsertColumns(" ", 0, 1));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.InsertColumns("Budget", 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.DeleteColumns(" ", new ExcelColumnSpan(0, 0)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.ShiftColumns(" ", new ExcelColumnSpan(0, 0), 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetColumnVisibility(" ", new ExcelColumnSpan(0, 0), true));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.GroupColumns(" ", new ExcelColumnSpan(0, 0), false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.UngroupColumns(" ", new ExcelColumnSpan(0, 0)));
  }

  @Test
  void validatesMergeAndLayoutSizingCommandInputs() {
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.MergeCells(null, "A1:B2"));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.MergeCells(" ", "A1"));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.MergeCells("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.MergeCells("Budget", " "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.UnmergeCells("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.UnmergeCells(" ", "A1:B2"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.UnmergeCells("Budget", " "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetColumnWidth(null, 0, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetColumnWidth(" ", 0, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetColumnWidth("Budget", -1, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetColumnWidth("Budget", 0, -1, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetColumnWidth("Budget", 2, 1, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetColumnWidth("Budget", 0, 0, 256.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetColumnWidth("Budget", 0, 0, Double.MIN_VALUE));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetColumnWidth("Budget", 0, 0, Double.POSITIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRowHeight("Budget", -1, 0, 28.5));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetRowHeight(null, 0, 0, 28.5));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetRowHeight(" ", 0, 0, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRowHeight("Budget", 0, -1, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRowHeight("Budget", 2, 1, 28.5));
    assertDoesNotThrow(
        () -> new WorkbookCommand.SetRowHeight("Budget", 0, 0, Short.MAX_VALUE / 20.0d));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCommand.SetRowHeight("Budget", 0, 0, Math.nextUp(Short.MAX_VALUE / 20.0d)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRowHeight("Budget", 0, 0, (Short.MAX_VALUE / 20.0d) + 1.0d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRowHeight("Budget", 0, 0, Double.MIN_VALUE));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRowHeight("Budget", 0, 0, 0.0));
  }

  @Test
  void validatesPaneZoomAndPrintLayoutCommandInputs() {
    assertDoesNotThrow(() -> new WorkbookCommand.SetSheetPane("Budget", new ExcelSheetPane.None()));
    assertDoesNotThrow(
        () -> new WorkbookCommand.SetSheetPane("Budget", new ExcelSheetPane.Frozen(0, 2, 0, 2)));
    assertDoesNotThrow(
        () ->
            new WorkbookCommand.SetSheetPane(
                "Budget", new ExcelSheetPane.Split(1200, 0, 3, 0, ExcelPaneRegion.UPPER_RIGHT)));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetSheetPane(null, new ExcelSheetPane.None()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetSheetPane(" ", new ExcelSheetPane.None()));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetSheetPane("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetSheetZoom("Budget", 9));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetSheetZoom("Budget", 401));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.SetSheetZoom(" ", 100));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetPrintLayout("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetPrintLayout(" ", defaultPrintLayout()));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.ClearPrintLayout(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.ClearPrintLayout(" "));
  }

  @Test
  void validatesCellRangeAndLinkCommentCommandInputs() {
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetCell(null, "A1", ExcelCellValue.text("x")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetCell("Budget", null, ExcelCellValue.text("x")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetCell("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetCell(" ", "A1", ExcelCellValue.text("x")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetCell("Budget", " ", ExcelCellValue.text("x")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetRange(null, "A1", List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetRange(" ", "A1", List.of()));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetRange("Budget", null, List.of()));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetRange("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRange("Budget", " ", List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRange("Budget", "A1", List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRange("Budget", "A1:B2", List.of(List.of())));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCommand.SetRange(
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
        () -> new WorkbookCommand.SetRange("Budget", "A1", rowsWithNull));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.ClearRange(null, "A1"));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.ClearRange("Budget", null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.ClearRange(" ", "A1"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.ClearRange("Budget", " "));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookCommand.SetHyperlink(
                null, "A1", new ExcelHyperlink.Url("https://example.com")));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookCommand.SetHyperlink(
                "Budget", null, new ExcelHyperlink.Url("https://example.com")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetHyperlink("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCommand.SetHyperlink(
                " ", "A1", new ExcelHyperlink.Url("https://example.com")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCommand.SetHyperlink(
                "Budget", " ", new ExcelHyperlink.Url("https://example.com")));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.ClearHyperlink(null, "A1"));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.ClearHyperlink("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.ClearHyperlink(" ", "A1"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.ClearHyperlink("Budget", " "));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookCommand.SetComment(
                null, "A1", new ExcelComment("Review", "GridGrind", false)));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookCommand.SetComment(
                "Budget", null, new ExcelComment("Review", "GridGrind", false)));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetComment("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCommand.SetComment(
                " ", "A1", new ExcelComment("Review", "GridGrind", false)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCommand.SetComment(
                "Budget", " ", new ExcelComment("Review", "GridGrind", false)));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.ClearComment(null, "A1"));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.ClearComment("Budget", null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.ClearComment(" ", "A1"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.ClearComment("Budget", " "));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.ApplyStyle(null, "A1", ExcelCellStyle.numberFormat("0")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.ApplyStyle("Budget", null, ExcelCellStyle.numberFormat("0")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.ApplyStyle("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.ApplyStyle(" ", "A1", ExcelCellStyle.numberFormat("0")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.ApplyStyle("Budget", " ", ExcelCellStyle.numberFormat("0")));
  }

  @Test
  void validatesValidationNamedRangeAndAppendCommandInputs() {
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetDataValidation("Budget", "A1", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetDataValidation(null, "A1", validationDefinition()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetDataValidation("Budget", null, validationDefinition()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetDataValidation(" ", "A1", validationDefinition()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetDataValidation("Budget", " ", validationDefinition()));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.ClearDataValidations("Budget", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.ClearDataValidations(null, new ExcelRangeSelection.All()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.ClearDataValidations(" ", new ExcelRangeSelection.All()));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookCommand.SetConditionalFormatting(
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
        () -> new WorkbookCommand.SetConditionalFormatting("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCommand.SetConditionalFormatting(
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
        () -> new WorkbookCommand.ClearConditionalFormatting("Budget", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.ClearConditionalFormatting(null, new ExcelRangeSelection.All()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.ClearConditionalFormatting(" ", new ExcelRangeSelection.All()));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetAutofilter(null, "A1:B2"));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetAutofilter("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetAutofilter(" ", "A1:B2"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetAutofilter("Budget", " "));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.ClearAutofilter(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.ClearAutofilter(" "));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.SetTable(null));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.DeleteTable(null, "Budget"));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.DeleteTable("BudgetTable", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.DeleteTable(" ", "Budget"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.DeleteTable("BudgetTable", " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCommand.SetTable(
                new ExcelTableDefinition(
                    "A1", "Budget", "A1:B2", false, new ExcelTableStyle.None())));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.SetNamedRange(null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.DeleteNamedRange(null, new ExcelNamedRangeScope.WorkbookScope()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.DeleteNamedRange("BudgetTotal", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCommand.DeleteNamedRange(
                "_XLNM.PRINT_AREA", new ExcelNamedRangeScope.WorkbookScope()));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.DeleteTable("A1", "Budget"));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.AppendRow(null, List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.AppendRow(" ", List.of()));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.AppendRow("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.AppendRow("Budget", List.of()));
    List<ExcelCellValue> valuesWithNull = new ArrayList<>();
    valuesWithNull.add(null);
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.AppendRow("Budget", valuesWithNull));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.AutoSizeColumns(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.AutoSizeColumns(" "));
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
      WorkbookCommand.CreateSheet createSheet,
      WorkbookCommand.RenameSheet renameSheet,
      WorkbookCommand.DeleteSheet deleteSheet,
      WorkbookCommand.MoveSheet moveSheet,
      WorkbookCommand.CopySheet copySheet,
      WorkbookCommand.SetActiveSheet setActiveSheet,
      WorkbookCommand.SetSelectedSheets setSelectedSheets,
      WorkbookCommand.SetSheetVisibility setSheetVisibility,
      WorkbookCommand.SetSheetProtection setSheetProtection,
      WorkbookCommand.ClearSheetProtection clearSheetProtection,
      WorkbookCommand.MergeCells mergeCells,
      WorkbookCommand.UnmergeCells unmergeCells,
      WorkbookCommand.SetColumnWidth setColumnWidth,
      WorkbookCommand.SetRowHeight setRowHeight,
      WorkbookCommand.SetSheetPane setSheetPane,
      WorkbookCommand.SetSheetZoom setSheetZoom,
      WorkbookCommand.SetPrintLayout setPrintLayout,
      WorkbookCommand.ClearPrintLayout clearPrintLayout,
      WorkbookCommand.SetCell setCell,
      WorkbookCommand.SetRange setRange,
      WorkbookCommand.ClearRange clearRange,
      WorkbookCommand.SetHyperlink setHyperlink,
      WorkbookCommand.ClearHyperlink clearHyperlink,
      WorkbookCommand.SetComment setComment,
      WorkbookCommand.ClearComment clearComment,
      WorkbookCommand.ApplyStyle applyStyle,
      WorkbookCommand.SetDataValidation setDataValidation,
      WorkbookCommand.ClearDataValidations clearDataValidations,
      WorkbookCommand.SetConditionalFormatting setConditionalFormatting,
      WorkbookCommand.ClearConditionalFormatting clearConditionalFormatting,
      WorkbookCommand.SetAutofilter setAutofilter,
      WorkbookCommand.ClearAutofilter clearAutofilter,
      WorkbookCommand.SetTable setTable,
      WorkbookCommand.DeleteTable deleteTable,
      WorkbookCommand.SetPivotTable setPivotTable,
      WorkbookCommand.DeletePivotTable deletePivotTable,
      WorkbookCommand.SetNamedRange setNamedRange,
      WorkbookCommand.DeleteNamedRange deleteNamedRange,
      WorkbookCommand.AppendRow appendRow,
      WorkbookCommand.AutoSizeColumns autoSizeColumns,
      WorkbookCommand.EvaluateAllFormulas evaluate,
      WorkbookCommand.EvaluateFormulaCells evaluateFormulaCells,
      WorkbookCommand.ClearFormulaCaches clearFormulaCaches,
      WorkbookCommand.ForceFormulaRecalculationOnOpen recalc) {}

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
