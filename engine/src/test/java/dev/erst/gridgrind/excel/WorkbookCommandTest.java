package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.layout.SheetPresentationSpec;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.protection.ResolvedWorkbookProtection;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelRowSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookCommand sealed interface record construction. */
class WorkbookCommandTest {
  @Test
  void createsSupportedCommandsAndCopiesCollections() {
    List<ResolvedCellWrite> values = new ArrayList<>(List.of(new CellValue.Text("Item")));
    List<List<ResolvedCellWrite>> rows =
        new ArrayList<>(
            List.of(
                new ArrayList<>(
                    List.of(new CellValue.Text("Item"), new CellValue.NumberValue(49.0))),
                new ArrayList<>(
                    List.of(new CellValue.Text("Tax"), new CellValue.NumberValue(10.0)))));
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
    List<AutofilterFilterColumnSpec> criteria =
        new ArrayList<>(
            List.of(
                new AutofilterFilterColumnSpec(
                    0L, false, new AutofilterFilterCriterionSpec.Values(List.of("Ada"), true))));
    WorkbookSheetCommand.SetSheetProtection sheetProtection =
        new WorkbookSheetCommand.SetSheetProtection(
            "Budget", protectionSettings(), Optional.of("gridgrind"));
    WorkbookSheetCommand.SetWorkbookProtection workbookProtection =
        new WorkbookSheetCommand.SetWorkbookProtection(
            new ResolvedWorkbookProtection(
                true, false, true, Optional.of("book"), Optional.of("review")));
    WorkbookSheetCommand.ClearWorkbookProtection clearWorkbookProtection =
        new WorkbookSheetCommand.ClearWorkbookProtection();
    WorkbookTabularCommand.SetAutofilter setAutofilter =
        new WorkbookTabularCommand.SetAutofilter(
            "Budget",
            "A1:C4",
            criteria,
            Optional.of(
                new AutofilterSortStateSpec(
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:C4"),
                    true,
                    false,
                    Optional.empty(),
                    List.of(
                        new AutofilterSortConditionSpec.Value(
                            new dev.erst.gridgrind.core.vocabulary.RangeRef("A2:A4"), false)))));

    criteria.clear();

    assertEquals("gridgrind", sheetProtection.password().orElseThrow());
    assertEquals("book", workbookProtection.protection().workbookPassword().orElseThrow());
    assertNotNull(clearWorkbookProtection);
    assertEquals(1, setAutofilter.criteria().size());
    assertFalse(setAutofilter.criteria().getFirst().showButton());
    assertTrue(setAutofilter.sortState().orElseThrow().caseSensitive());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSheetCommand.SetSheetProtection(
                "Budget", protectionSettings(), Optional.of(" ")));
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
      List<ResolvedCellWrite> values, List<List<ResolvedCellWrite>> rows, CellStylePatch style) {
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
            "Budget", "A1", new CellValueSpec.Date(LocalDate.of(2026, 3, 23))),
        new WorkbookCellCommand.SetRange("Budget", "A1:B2", rows),
        new WorkbookCellCommand.ClearRange("Budget", "C1:C2"),
        new WorkbookAnnotationCommand.SetHyperlink(
            "Budget", "A1", new HyperlinkTarget.Url("https://example.com/report")),
        new WorkbookAnnotationCommand.ClearHyperlink("Budget", "A1"),
        new WorkbookAnnotationCommand.SetComment(
            "Budget", "A1", new ResolvedComment("Review", "GridGrind", false)),
        new WorkbookAnnotationCommand.ClearComment("Budget", "A1"),
        new WorkbookFormattingCommand.ApplyStyle("Budget", "A1:B1", style),
        new WorkbookFormattingCommand.SetDataValidation("Budget", "B2:B5", validationDefinition()),
        new WorkbookFormattingCommand.ClearDataValidations(
            "Budget", new ExcelRangeSelection.Selected(List.of("C2:D4"))),
        new WorkbookFormattingCommand.SetConditionalFormatting(
            "Budget",
            List.of("A2:A5").stream()
                .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                .toList(),
            new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
                List.of(
                    new ConditionalFormattingRuleSpec.FormulaRule(
                        "A2>0",
                        true,
                        Optional.of(
                            new DifferentialStylePatch(
                                Optional.of("0.00"),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty())))))),
        new WorkbookFormattingCommand.ClearConditionalFormatting(
            "Budget", new ExcelRangeSelection.Selected(List.of("A2:A5"))),
        new WorkbookTabularCommand.SetAutofilter("Budget", "A1:C4"),
        new WorkbookTabularCommand.ClearAutofilter("Budget"),
        new WorkbookTabularCommand.SetTable(
            new ResolvedTable(
                new TableName("BudgetTable"),
                new SheetName("Budget"),
                new RangeRef("A1:C4"),
                true,
                new TableStyleSpec.Named("TableStyleMedium2", false, false, true, false))),
        new WorkbookTabularCommand.DeleteTable("BudgetTable", "Budget"),
        new WorkbookTabularCommand.SetPivotTable(
            new PivotTableSpec(
                "Budget Pivot",
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:C4")),
                new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress("E3")),
                List.of("Item"),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableSpec.DataField(
                        "Tax",
                        ExcelPivotDataConsolidateFunction.SUM,
                        "Total Tax",
                        Optional.of("#,##0.00"))))),
        new WorkbookTabularCommand.DeletePivotTable("Budget Pivot", "Budget"),
        new WorkbookMetadataCommand.SetNamedRange(
            new StructuredMutationAction.SetNamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetTotal"),
                new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("B4")))),
        new WorkbookMetadataCommand.DeleteNamedRange(
            "BudgetTotal",
            new NamedRangeScopeReport.Sheet(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"))),
        new WorkbookCellCommand.AppendRow("Budget", values),
        new WorkbookLayoutCommand.AutoSizeColumns("Budget"));
  }

  private void assertCreatedCommands(CreatedCommands commands, CellStylePatch style) {
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
    assertInstanceOf(HyperlinkTarget.Url.class, commands.setHyperlink().target());
    assertEquals("A1", commands.clearHyperlink().address());
    assertEquals("Review", commands.setComment().comment().text());
    assertEquals("A1", commands.clearComment().address());
    assertEquals(style, commands.applyStyle().style());
    assertEquals("B2:B5", commands.setDataValidation().range());
    assertEquals(
        List.of("C2:D4"),
        ((ExcelRangeSelection.Selected) commands.clearDataValidations().selection()).ranges());
    assertEquals("Budget", commands.setConditionalFormatting().sheetName());
    assertEquals(
        List.of("A2:A5"),
        commands.setConditionalFormatting().ranges().stream()
            .map(dev.erst.gridgrind.core.vocabulary.RangeRef::value)
            .toList());
    assertEquals(
        List.of("A2:A5"),
        ((ExcelRangeSelection.Selected) commands.clearConditionalFormatting().selection())
            .ranges());
    assertEquals("A1:C4", commands.setAutofilter().range());
    assertEquals("Budget", commands.clearAutofilter().sheetName());
    assertEquals("BudgetTable", commands.setTable().definition().name().value());
    assertEquals("Budget", commands.deleteTable().sheetName());
    assertEquals("Budget Pivot", commands.setPivotTable().definition().name());
    assertEquals("Budget", commands.deletePivotTable().sheetName());
    assertEquals("BudgetTotal", commands.setNamedRange().definition().name().value());
    assertEquals(
        "Budget",
        ((NamedRangeScopeReport.Sheet) commands.deleteNamedRange().scope()).sheetName().value());
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
                    "Budget", SheetPresentationSpec.defaults()));
    assertEquals("Budget", setSheetPresentation.sheetName());
    assertEquals(SheetPresentationSpec.defaults(), setSheetPresentation.presentation());
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookLayoutCommand.SetSheetPresentation(null, SheetPresentationSpec.defaults()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookLayoutCommand.SetSheetPresentation(" ", SheetPresentationSpec.defaults()));
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
        () -> new WorkbookCellCommand.SetCell(null, "A1", new CellValue.Text("x")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCellCommand.SetCell("Budget", null, new CellValue.Text("x")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCellCommand.SetCell("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCellCommand.SetCell(" ", "A1", new CellValue.Text("x")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCellCommand.SetCell("Budget", " ", new CellValue.Text("x")));
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
                    List.of(new CellValue.Text("x")),
                    List.of(new CellValue.Text("y"), new CellValue.Text("z")))));
    List<List<ResolvedCellWrite>> rowsWithNull = new ArrayList<>();
    List<ResolvedCellWrite> rowWithNull = new ArrayList<>();
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
                null, "A1", new HyperlinkTarget.Url("https://example.com")));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookAnnotationCommand.SetHyperlink(
                "Budget", null, new HyperlinkTarget.Url("https://example.com")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnnotationCommand.SetHyperlink("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookAnnotationCommand.SetHyperlink(
                " ", "A1", new HyperlinkTarget.Url("https://example.com")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookAnnotationCommand.SetHyperlink(
                "Budget", " ", new HyperlinkTarget.Url("https://example.com")));
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
                null, "A1", new ResolvedComment("Review", "GridGrind", false)));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookAnnotationCommand.SetComment(
                "Budget", null, new ResolvedComment("Review", "GridGrind", false)));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnnotationCommand.SetComment("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookAnnotationCommand.SetComment(
                " ", "A1", new ResolvedComment("Review", "GridGrind", false)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookAnnotationCommand.SetComment(
                "Budget", " ", new ResolvedComment("Review", "GridGrind", false)));
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
            new WorkbookFormattingCommand.ApplyStyle(
                null,
                "A1",
                new CellStylePatch(
                    Optional.of("0"),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookFormattingCommand.ApplyStyle(
                "Budget",
                null,
                new CellStylePatch(
                    Optional.of("0"),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookFormattingCommand.ApplyStyle("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookFormattingCommand.ApplyStyle(
                " ",
                "A1",
                new CellStylePatch(
                    Optional.of("0"),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookFormattingCommand.ApplyStyle(
                "Budget",
                " ",
                new CellStylePatch(
                    Optional.of("0"),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())));
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
                List.of("A2:A5").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                    .toList(),
                new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
                    List.of(
                        new ConditionalFormattingRuleSpec.FormulaRule(
                            "A2>0",
                            true,
                            Optional.of(
                                new DifferentialStylePatch(
                                    Optional.of("0.00"),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty())))))));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookFormattingCommand.SetConditionalFormatting("Budget", List.of(), null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookFormattingCommand.SetConditionalFormatting(
                " ",
                List.of("A2:A5").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                    .toList(),
                new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
                    List.of(
                        new ConditionalFormattingRuleSpec.FormulaRule(
                            "A2>0",
                            true,
                            Optional.of(
                                new DifferentialStylePatch(
                                    Optional.of("0.00"),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty())))))));
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
                new ResolvedTable(
                    new TableName("A1"),
                    new SheetName("Budget"),
                    new RangeRef("A1:B2"),
                    false,
                    new TableStyleSpec.None())));
    assertThrows(NullPointerException.class, () -> new WorkbookMetadataCommand.SetNamedRange(null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookMetadataCommand.DeleteNamedRange(
                null, new NamedRangeScopeReport.Workbook()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookMetadataCommand.DeleteNamedRange("BudgetTotal", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookMetadataCommand.DeleteNamedRange(
                "_XLNM.PRINT_AREA", new NamedRangeScopeReport.Workbook()));
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
    List<ResolvedCellWrite> valuesWithNull = new ArrayList<>();
    valuesWithNull.add(null);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCellCommand.AppendRow("Budget", valuesWithNull));
    assertThrows(NullPointerException.class, () -> new WorkbookLayoutCommand.AutoSizeColumns(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookLayoutCommand.AutoSizeColumns(" "));
  }

  private static ResolvedDataValidation validationDefinition() {
    return new ResolvedDataValidation(
        new DataValidationRuleSpec.TextLength(
            ExcelComparisonOperator.LESS_OR_EQUAL,
            ExcelComparisonFormulaSupport.normalizeFormula("20", "formula1"),
            ExcelComparisonFormulaSupport.normalizeOptionalFormula(Optional.empty(), "formula2")),
        true,
        false,
        Optional.of(new ResolvedValidationPrompt("Reason", "Use 20 characters or fewer.", true)),
        Optional.of(
            new ResolvedValidationErrorAlert(
                ExcelDataValidationErrorStyle.STOP, "Too long", "Use a shorter reason.", true)));
  }

  private CellStylePatch defaultStyle() {
    return new CellStylePatch(
        Optional.of("#,##0.00"),
        Optional.of(
            new CellAlignmentPatch(
                Optional.of(true),
                Optional.of(ExcelHorizontalAlignment.RIGHT),
                Optional.of(ExcelVerticalAlignment.CENTER),
                Optional.empty(),
                Optional.empty())),
        Optional.of(
            new CellFontPatch(
                Optional.of(true),
                Optional.of(false),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty())),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
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

  private static ResolvedPrintLayout defaultPrintLayout() {
    return new ResolvedPrintLayout(
        new PrintAreaSpec.Range(new RangeRef("A1:C20")),
        ExcelPrintOrientation.LANDSCAPE,
        new PrintScalingSpec.Fit(1, 0),
        new PrintTitleRowsSpec.Band(0, 1),
        new PrintTitleColumnsSpec.Band(0, 0),
        new HeaderFooterTextReport("Left", "Center", "Right"),
        new HeaderFooterTextReport("Footer Left", "", "Footer Right"));
  }

  private static SheetProtectionSettings protectionSettings() {
    return new SheetProtectionSettings(
        true, false, true, false, true, false, true, false, true, false, true, false, true, false,
        true);
  }
}
