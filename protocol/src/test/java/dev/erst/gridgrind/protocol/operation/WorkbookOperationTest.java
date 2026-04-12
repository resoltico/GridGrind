package dev.erst.gridgrind.protocol.operation;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.ExcelColumnSpan;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelPaneRegion;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import dev.erst.gridgrind.excel.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.ExcelRowSpan;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.protocol.dto.*;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookOperation record construction and operationType behavior. */
class WorkbookOperationTest {
  @Test
  void buildsSheetAndLayoutOperations() {
    WorkbookOperation.EnsureSheet ensureSheet = new WorkbookOperation.EnsureSheet("Budget");
    WorkbookOperation.RenameSheet renameSheet =
        new WorkbookOperation.RenameSheet("Budget", "Summary");
    WorkbookOperation.DeleteSheet deleteSheet = new WorkbookOperation.DeleteSheet("Archive");
    WorkbookOperation.MoveSheet moveSheet = new WorkbookOperation.MoveSheet("Budget", 1);
    WorkbookOperation.MergeCells mergeCells = new WorkbookOperation.MergeCells("Budget", "A1:B2");
    WorkbookOperation.UnmergeCells unmergeCells =
        new WorkbookOperation.UnmergeCells("Budget", "A1:B2");
    WorkbookOperation.SetColumnWidth setColumnWidth =
        new WorkbookOperation.SetColumnWidth("Budget", 0, 2, 16.0);
    WorkbookOperation.SetRowHeight setRowHeight =
        new WorkbookOperation.SetRowHeight("Budget", 0, 3, 28.5);
    WorkbookOperation.SetSheetPane setSheetPane =
        new WorkbookOperation.SetSheetPane("Budget", new PaneInput.Frozen(1, 2, 1, 2));
    WorkbookOperation.SetSheetZoom setSheetZoom = new WorkbookOperation.SetSheetZoom("Budget", 135);
    WorkbookOperation.SetPrintLayout setPrintLayout =
        new WorkbookOperation.SetPrintLayout("Budget", defaultPrintLayout());
    WorkbookOperation.ClearPrintLayout clearPrintLayout =
        new WorkbookOperation.ClearPrintLayout("Budget");

    assertEquals("Budget", ensureSheet.sheetName());
    assertEquals("Summary", renameSheet.newSheetName());
    assertEquals("Archive", deleteSheet.sheetName());
    assertEquals(1, moveSheet.targetIndex());
    assertEquals("A1:B2", mergeCells.range());
    assertEquals("A1:B2", unmergeCells.range());
    assertEquals(16.0, setColumnWidth.widthCharacters());
    assertEquals(28.5, setRowHeight.heightPoints());
    assertEquals(new PaneInput.Frozen(1, 2, 1, 2), setSheetPane.pane());
    assertEquals(135, setSheetZoom.zoomPercent());
    assertEquals(defaultPrintLayout(), setPrintLayout.printLayout());
    assertEquals("Budget", clearPrintLayout.sheetName());
  }

  @Test
  void buildsRowAndColumnStructureOperations() {
    WorkbookOperation.InsertRows insertRows = new WorkbookOperation.InsertRows("Budget", 2, 3);
    WorkbookOperation.DeleteRows deleteRows =
        new WorkbookOperation.DeleteRows("Budget", new RowSpanInput.Band(4, 6));
    WorkbookOperation.ShiftRows shiftRows =
        new WorkbookOperation.ShiftRows("Budget", new RowSpanInput.Band(1, 3), 2);
    WorkbookOperation.InsertColumns insertColumns =
        new WorkbookOperation.InsertColumns("Budget", 1, 2);
    WorkbookOperation.DeleteColumns deleteColumns =
        new WorkbookOperation.DeleteColumns("Budget", new ColumnSpanInput.Band(3, 4));
    WorkbookOperation.ShiftColumns shiftColumns =
        new WorkbookOperation.ShiftColumns("Budget", new ColumnSpanInput.Band(0, 1), -1);
    WorkbookOperation.SetRowVisibility setRowVisibility =
        new WorkbookOperation.SetRowVisibility("Budget", new RowSpanInput.Band(5, 7), true);
    WorkbookOperation.SetColumnVisibility setColumnVisibility =
        new WorkbookOperation.SetColumnVisibility("Budget", new ColumnSpanInput.Band(2, 3), false);
    WorkbookOperation.GroupRows groupRows =
        new WorkbookOperation.GroupRows("Budget", new RowSpanInput.Band(8, 10), null);
    WorkbookOperation.UngroupRows ungroupRows =
        new WorkbookOperation.UngroupRows("Budget", new RowSpanInput.Band(8, 10));
    WorkbookOperation.GroupColumns groupColumns =
        new WorkbookOperation.GroupColumns("Budget", new ColumnSpanInput.Band(4, 6), true);
    WorkbookOperation.GroupColumns groupColumnsWithDefaultCollapse =
        new WorkbookOperation.GroupColumns("Budget", new ColumnSpanInput.Band(4, 6), null);
    WorkbookOperation.UngroupColumns ungroupColumns =
        new WorkbookOperation.UngroupColumns("Budget", new ColumnSpanInput.Band(4, 6));

    assertEquals(2, insertRows.rowIndex());
    assertEquals(3, insertRows.rowCount());
    assertEquals(new RowSpanInput.Band(4, 6), deleteRows.rows());
    assertEquals(2, shiftRows.delta());
    assertEquals(1, insertColumns.columnIndex());
    assertEquals(new ColumnSpanInput.Band(3, 4), deleteColumns.columns());
    assertEquals(-1, shiftColumns.delta());
    assertTrue(setRowVisibility.hidden());
    assertFalse(setColumnVisibility.hidden());
    assertFalse(groupRows.collapsed());
    assertEquals(new RowSpanInput.Band(8, 10), ungroupRows.rows());
    assertTrue(groupColumns.collapsed());
    assertFalse(groupColumnsWithDefaultCollapse.collapsed());
    assertEquals(new ColumnSpanInput.Band(4, 6), ungroupColumns.columns());
  }

  @Test
  void buildsSheetStateOperationsAndDefaultsCopyPosition() {
    SheetProtectionSettings protection = protectionSettings();
    WorkbookOperation.CopySheet copySheet =
        new WorkbookOperation.CopySheet("Budget", "Budget Copy", null);
    WorkbookOperation.SetActiveSheet setActiveSheet =
        new WorkbookOperation.SetActiveSheet("Budget Copy");
    WorkbookOperation.SetSelectedSheets setSelectedSheets =
        new WorkbookOperation.SetSelectedSheets(List.of("Budget", "Budget Copy"));
    WorkbookOperation.SetSheetVisibility setSheetVisibility =
        new WorkbookOperation.SetSheetVisibility("Budget", ExcelSheetVisibility.HIDDEN);
    WorkbookOperation.SetSheetProtection setSheetProtection =
        new WorkbookOperation.SetSheetProtection("Budget", protection);
    WorkbookOperation.ClearSheetProtection clearSheetProtection =
        new WorkbookOperation.ClearSheetProtection("Budget");

    assertEquals("Budget", copySheet.sourceSheetName());
    assertEquals("Budget Copy", copySheet.newSheetName());
    assertInstanceOf(SheetCopyPosition.AppendAtEnd.class, copySheet.position());
    assertEquals("Budget Copy", setActiveSheet.sheetName());
    assertEquals(List.of("Budget", "Budget Copy"), setSelectedSheets.sheetNames());
    assertEquals(ExcelSheetVisibility.HIDDEN, setSheetVisibility.visibility());
    assertEquals(protection, setSheetProtection.protection());
    assertEquals("Budget", clearSheetProtection.sheetName());
  }

  @Test
  void buildsCellAndMetadataOperationsAndCopiesCollections() {
    List<List<CellInput>> rows = protocolRows();
    CellStyleInput style = protocolStyle();
    WorkbookOperation.SetCell setCell =
        new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("Item"));
    WorkbookOperation.SetRange setRange = new WorkbookOperation.SetRange("Budget", "A1:B2", rows);
    WorkbookOperation.ClearRange clearRange = new WorkbookOperation.ClearRange("Budget", "C1:C4");
    WorkbookOperation.SetHyperlink setHyperlink =
        new WorkbookOperation.SetHyperlink(
            "Budget", "A1", new HyperlinkTarget.Url("https://example.com/report"));
    WorkbookOperation.ClearHyperlink clearHyperlink =
        new WorkbookOperation.ClearHyperlink("Budget", "A1");
    WorkbookOperation.SetComment setComment =
        new WorkbookOperation.SetComment(
            "Budget", "A1", new CommentInput("Review", "GridGrind", null));
    WorkbookOperation.ClearComment clearComment =
        new WorkbookOperation.ClearComment("Budget", "A1");
    WorkbookOperation.ApplyStyle applyStyle =
        new WorkbookOperation.ApplyStyle("Budget", "B1:B2", style);

    rows.clear();

    assertEquals("A1", setCell.address());
    assertEquals("A1:B2", setRange.range());
    assertEquals(2, setRange.rows().size());
    assertEquals("C1:C4", clearRange.range());
    assertEquals(new HyperlinkTarget.Url("https://example.com/report"), setHyperlink.target());
    assertEquals("A1", clearHyperlink.address());
    assertFalse(setComment.comment().visible());
    assertEquals("A1", clearComment.address());
    assertEquals(style, applyStyle.style());
  }

  @Test
  void buildsDrawingOperationsAndUsesPhase5OperationTypes() {
    DrawingAnchorInput.TwoCell anchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(1, 2, 3, 4),
            new DrawingMarkerInput(4, 6, 7, 8),
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    PictureDataInput pictureData =
        new PictureDataInput(
            ExcelPictureFormat.PNG,
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");
    WorkbookOperation.SetPicture setPicture =
        new WorkbookOperation.SetPicture(
            "Budget", new PictureInput("OpsPicture", pictureData, anchor, "Queue preview"));
    WorkbookOperation.SetShape setShape =
        new WorkbookOperation.SetShape(
            "Budget",
            new ShapeInput(
                "OpsShape", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, anchor, "rect", "Queue"));
    WorkbookOperation.SetEmbeddedObject setEmbeddedObject =
        new WorkbookOperation.SetEmbeddedObject(
            "Budget",
            new EmbeddedObjectInput(
                "OpsEmbed",
                "Payload",
                "payload.txt",
                "payload.txt",
                "cGF5bG9hZA==",
                pictureData,
                anchor));
    WorkbookOperation.SetDrawingObjectAnchor setDrawingObjectAnchor =
        new WorkbookOperation.SetDrawingObjectAnchor("Budget", "OpsPicture", anchor);
    WorkbookOperation.DeleteDrawingObject deleteDrawingObject =
        new WorkbookOperation.DeleteDrawingObject("Budget", "OpsPicture");

    assertEquals("OpsPicture", setPicture.picture().name());
    assertEquals("OpsShape", setShape.shape().name());
    assertEquals("OpsEmbed", setEmbeddedObject.embeddedObject().name());
    assertEquals("OpsPicture", setDrawingObjectAnchor.objectName());
    assertEquals("OpsPicture", deleteDrawingObject.objectName());
    assertEquals("SET_PICTURE", setPicture.operationType());
    assertEquals("SET_SHAPE", setShape.operationType());
    assertEquals("SET_EMBEDDED_OBJECT", setEmbeddedObject.operationType());
    assertEquals("SET_DRAWING_OBJECT_ANCHOR", setDrawingObjectAnchor.operationType());
    assertEquals("DELETE_DRAWING_OBJECT", deleteDrawingObject.operationType());
  }

  @Test
  void buildsValidationNamedRangeAndTerminalOperationsAndCopiesCollections() {
    List<CellInput> rowValues = protocolRowValues();
    WorkbookOperation.SetDataValidation setDataValidation =
        new WorkbookOperation.SetDataValidation("Budget", "B2:B5", protocolValidation());
    WorkbookOperation.ClearDataValidations clearDataValidations =
        new WorkbookOperation.ClearDataValidations(
            "Budget", new RangeSelection.Selected(List.of("C2:D4")));
    WorkbookOperation.SetConditionalFormatting setConditionalFormatting =
        new WorkbookOperation.SetConditionalFormatting(
            "Budget",
            new ConditionalFormattingBlockInput(
                List.of("E2:E5"),
                List.of(
                    new ConditionalFormattingRuleInput.FormulaRule(
                        "E2>0",
                        true,
                        new DifferentialStyleInput(
                            null, true, null, null, "#112233", null, null, null, null)))));
    WorkbookOperation.ClearConditionalFormatting clearConditionalFormatting =
        new WorkbookOperation.ClearConditionalFormatting(
            "Budget", new RangeSelection.Selected(List.of("E2:E5")));
    WorkbookOperation.SetAutofilter setAutofilter =
        new WorkbookOperation.SetAutofilter("Budget", "A1:C4");
    WorkbookOperation.ClearAutofilter clearAutofilter =
        new WorkbookOperation.ClearAutofilter("Budget");
    WorkbookOperation.SetTable setTable =
        new WorkbookOperation.SetTable(
            new TableInput(
                "BudgetTable",
                "Budget",
                "A1:C4",
                true,
                new TableStyleInput.Named("TableStyleMedium2", false, false, true, false)));
    WorkbookOperation.DeleteTable deleteTable =
        new WorkbookOperation.DeleteTable("BudgetTable", "Budget");
    WorkbookOperation.SetNamedRange setNamedRange =
        new WorkbookOperation.SetNamedRange(
            "BudgetTotal", new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "B4"));
    WorkbookOperation.DeleteNamedRange deleteNamedRange =
        new WorkbookOperation.DeleteNamedRange("BudgetTotal", new NamedRangeScope.Sheet("Budget"));
    WorkbookOperation.AppendRow appendRow = new WorkbookOperation.AppendRow("Budget", rowValues);
    WorkbookOperation.AutoSizeColumns autoSizeColumns =
        new WorkbookOperation.AutoSizeColumns("Budget");
    WorkbookOperation.EvaluateFormulas evaluateFormulas = new WorkbookOperation.EvaluateFormulas();
    WorkbookOperation.EvaluateFormulaCells evaluateFormulaCells =
        new WorkbookOperation.EvaluateFormulaCells(
            List.of(new FormulaCellTargetInput("Budget", "B4")));
    WorkbookOperation.ClearFormulaCaches clearFormulaCaches =
        new WorkbookOperation.ClearFormulaCaches();
    WorkbookOperation.ForceFormulaRecalculationOnOpen recalcOnOpen =
        new WorkbookOperation.ForceFormulaRecalculationOnOpen();

    rowValues.clear();

    assertEquals("B2:B5", setDataValidation.range());
    assertInstanceOf(
        DataValidationRuleInput.TextLength.class, setDataValidation.validation().rule());
    assertEquals(
        List.of("C2:D4"), ((RangeSelection.Selected) clearDataValidations.selection()).ranges());
    assertEquals(List.of("E2:E5"), setConditionalFormatting.conditionalFormatting().ranges());
    assertEquals(
        List.of("E2:E5"),
        ((RangeSelection.Selected) clearConditionalFormatting.selection()).ranges());
    assertEquals("A1:C4", setAutofilter.range());
    assertEquals("Budget", clearAutofilter.sheetName());
    assertEquals("BudgetTable", setTable.table().name());
    assertEquals("Budget", deleteTable.sheetName());
    assertEquals("BudgetTotal", setNamedRange.name());
    assertEquals("Budget", ((NamedRangeScope.Sheet) deleteNamedRange.scope()).sheetName());
    assertEquals(1, appendRow.values().size());
    assertEquals("Budget", autoSizeColumns.sheetName());
    assertEquals("SET_DATA_VALIDATION", setDataValidation.operationType());
    assertEquals("CLEAR_DATA_VALIDATIONS", clearDataValidations.operationType());
    assertEquals("SET_CONDITIONAL_FORMATTING", setConditionalFormatting.operationType());
    assertEquals("CLEAR_CONDITIONAL_FORMATTING", clearConditionalFormatting.operationType());
    assertEquals("SET_AUTOFILTER", setAutofilter.operationType());
    assertEquals("CLEAR_AUTOFILTER", clearAutofilter.operationType());
    assertEquals("SET_TABLE", setTable.operationType());
    assertEquals("DELETE_TABLE", deleteTable.operationType());
    assertEquals("EVALUATE_FORMULAS", evaluateFormulas.operationType());
    assertEquals(List.of(new FormulaCellTargetInput("Budget", "B4")), evaluateFormulaCells.cells());
    assertEquals("EVALUATE_FORMULA_CELLS", evaluateFormulaCells.operationType());
    assertEquals("CLEAR_FORMULA_CACHES", clearFormulaCaches.operationType());
    assertEquals("FORCE_FORMULA_RECALCULATION_ON_OPEN", recalcOnOpen.operationType());
  }

  @Test
  @SuppressWarnings("PMD.NcssCount")
  void validatesNullAndEmptyCollectionConstraints() {
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.CopySheet(null, "Budget Copy", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.CopySheet("Budget", null, new SheetCopyPosition.AppendAtEnd()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.CopySheet("Budget", " ", new SheetCopyPosition.AppendAtEnd()));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.SetActiveSheet(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookOperation.SetActiveSheet(" "));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.SetSelectedSheets(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.SetSelectedSheets(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetSelectedSheets(List.of("Budget", "Budget")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetSheetVisibility("Budget", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetSheetProtection("Budget", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.ClearSheetProtection(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.AppendRow("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.EvaluateFormulaCells(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.EvaluateFormulaCells(
                List.of(
                    new FormulaCellTargetInput("Budget", "B4"),
                    new FormulaCellTargetInput("Budget", "B4"))));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.AutoSizeColumns(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookOperation.AutoSizeColumns(" "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.RenameSheet(null, "Summary"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.RenameSheet(" ", "Summary"));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.RenameSheet("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.RenameSheet("Budget", " "));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.DeleteSheet(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookOperation.DeleteSheet(" "));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.MoveSheet("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.MoveSheet("Budget", -1));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.MergeCells(null, "A1:B2"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.MergeCells("Budget", " "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.UnmergeCells("Budget", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetColumnWidth(null, 0, 0, 16.0));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetColumnWidth("Budget", null, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetColumnWidth("Budget", 1, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetColumnWidth("Budget", 0, 0, 0.0));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetRowHeight("Budget", null, 0, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetRowHeight("Budget", 2, 1, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetRowHeight("Budget", 0, 0, Double.NaN));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.InsertRows("Budget", null, 1));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.InsertRows("Budget", 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.InsertRows("Budget", ExcelRowSpan.MAX_ROW_INDEX + 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.DeleteRows("Budget", new RowSpanInput.Band(2, 1)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.ShiftRows("Budget", new RowSpanInput.Band(0, 0), 0));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.InsertColumns("Budget", 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.InsertColumns("Budget", ExcelColumnSpan.MAX_COLUMN_INDEX + 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.DeleteColumns("Budget", new ColumnSpanInput.Band(4, 3)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.ShiftColumns("Budget", new ColumnSpanInput.Band(0, 0), 0));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetRowVisibility("Budget", null, true));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetColumnVisibility("Budget", null, true));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.GroupRows("Budget", null, false));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.GroupColumns("Budget", null, false));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetSheetPane("Budget", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetSheetZoom("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.SetSheetZoom("Budget", 9));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.SetSheetZoom("Budget", 401));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetPrintLayout("Budget", null));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.ClearPrintLayout(null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetHyperlink("Budget", "A1", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetComment("Budget", "A1", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetPicture("Budget", null));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.SetShape("Budget", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetEmbeddedObject("Budget", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetDrawingObjectAnchor("Budget", "OpsPicture", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.DeleteDrawingObject("Budget", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetDataValidation("Budget", "A1", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.ClearDataValidations("Budget", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetConditionalFormatting("Budget", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.ClearConditionalFormatting("Budget", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetAutofilter(null, "A1:B2"));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.ClearAutofilter(null));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.SetTable(null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.DeleteTable(null, "Budget"));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.DeleteTable("BudgetTable", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookOperation.SetNamedRange(
                "BudgetTotal", null, new NamedRangeTarget("Budget", "B4")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.DeleteNamedRange("BudgetTotal", null));
  }

  @Test
  void validatesOperationRequirements() {
    assertThrows(IllegalArgumentException.class, () -> new WorkbookOperation.EnsureSheet(" "));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.RenameSheet("Budget", " "));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookOperation.DeleteSheet(" "));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.MoveSheet("Budget", -1));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.MergeCells("Budget", " "));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.UnmergeCells("Budget", " "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetColumnWidth("Budget", -1, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetRowHeight("Budget", 0, -1, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetSheetPane("Budget", new PaneInput.Frozen(1, 0, 0, 0)));

    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetCell("Budget", null, new CellInput.Text("x")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetCell("Budget", "A1", null));

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetRange("Budget", "A1:B2", List.of()));

    List<List<CellInput>> rowsWithNullRow = new ArrayList<>();
    rowsWithNullRow.add(null);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetRange("Budget", "A1:B1", rowsWithNullRow));

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetRange("Budget", "A1:B2", List.of(List.of())));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.SetRange(
                "Budget",
                "A1:B2",
                List.of(
                    List.of(new CellInput.Text("x")),
                    List.of(new CellInput.Text("y"), new CellInput.Text("z")))));

    List<List<CellInput>> rowsWithNullValue = new ArrayList<>();
    List<CellInput> rowWithNullValue = new ArrayList<>();
    rowWithNullValue.add(null);
    rowsWithNullValue.add(rowWithNullValue);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetRange("Budget", "A1", rowsWithNullValue));

    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.ApplyStyle("Budget", "A1:A2", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.SetHyperlink(
                "Budget", "A1", new HyperlinkTarget.Url("relative")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.SetComment(
                "Budget", "A1", new CommentInput(" ", "GridGrind", null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.SetDataValidation(
                "Budget",
                " ",
                new DataValidationInput(
                    new DataValidationRuleInput.CustomFormula("LEN(A1)>0"),
                    false,
                    false,
                    null,
                    null)));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.SetAutofilter("Budget", " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.SetTable(
                new TableInput("A1", "Budget", "A1:B2", false, new TableStyleInput.None())));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.DeleteTable("A1", "Budget"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.SetNamedRange(
                "A1", new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "B4")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.DeleteNamedRange(
                "_xlnm.Print_Area", new NamedRangeScope.Workbook()));

    List<CellInput> valuesWithNull = new ArrayList<>();
    valuesWithNull.add(null);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.AppendRow("Budget", valuesWithNull));

    // null rows list is coalesced to empty, which then fails the non-empty check
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.SetRange("Budget", "A1", null));
  }

  @Test
  void validatesColumnWidthHelperBranches() {
    assertDoesNotThrow(() -> WorkbookOperation.Validation.requireColumnWidthCharacters(8.43d));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookOperation.Validation.requireColumnWidthCharacters(256.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookOperation.Validation.requireColumnWidthCharacters(Double.MIN_VALUE));
  }

  @Test
  void validatesRowHeightHelperBranches() {
    assertDoesNotThrow(() -> WorkbookOperation.Validation.requireRowHeightPoints(15.0d));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            WorkbookOperation.Validation.requireRowHeightPoints((Short.MAX_VALUE / 20.0d) + 1.0d));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookOperation.Validation.requireRowHeightPoints(Double.MIN_VALUE));
  }

  @Test
  void validatesPaneAndZoomHelperBranches() {
    assertDoesNotThrow(() -> new PaneInput.Frozen(1, 2, 1, 2));
    assertDoesNotThrow(() -> new PaneInput.Frozen(0, 2, 0, 2));
    assertDoesNotThrow(() -> new PaneInput.Frozen(2, 0, 2, 0));
    assertDoesNotThrow(() -> new PaneInput.Split(1200, 0, 3, 0, ExcelPaneRegion.UPPER_RIGHT));
    assertDoesNotThrow(() -> WorkbookOperation.Validation.requireZoomPercent(10));
    assertDoesNotThrow(() -> WorkbookOperation.Validation.requireZoomPercent(400));
    assertThrows(IllegalArgumentException.class, () -> new PaneInput.Frozen(0, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new PaneInput.Frozen(0, 1, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new PaneInput.Frozen(1, 0, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new PaneInput.Frozen(2, 1, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new PaneInput.Frozen(1, 2, 1, 1));
    assertThrows(
        IllegalArgumentException.class, () -> WorkbookOperation.Validation.requireZoomPercent(9));
    assertThrows(
        IllegalArgumentException.class, () -> WorkbookOperation.Validation.requireZoomPercent(401));
  }

  @Test
  void rejectsSheetNamesExceeding31Characters() {
    String tooLong = "A".repeat(32);
    String exactly31 = "A".repeat(31);

    assertDoesNotThrow(() -> new WorkbookOperation.EnsureSheet(exactly31));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookOperation.EnsureSheet(tooLong));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.RenameSheet("Budget", tooLong));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.RenameSheet(tooLong, "NewName"));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookOperation.DeleteSheet(tooLong));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookOperation.MoveSheet(tooLong, 0));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.MergeCells(tooLong, "A1:B2"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetColumnWidth(tooLong, 0, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetRowHeight(tooLong, 0, 0, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetSheetPane(tooLong, new PaneInput.Frozen(1, 0, 1, 0)));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.SetSheetZoom(tooLong, 100));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetPrintLayout(tooLong, defaultPrintLayout()));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.ClearPrintLayout(tooLong));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetCell(tooLong, "A1", new CellInput.Text("x")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.SetRange(
                tooLong, "A1", List.of(List.of(new CellInput.Text("x")))));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.ClearRange(tooLong, "A1:B2"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.SetHyperlink(
                tooLong, "A1", new HyperlinkTarget.Url("https://example.com")));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.ClearHyperlink(tooLong, "A1"));
    CellStyleInput style =
        new CellStyleInput(
            null, new CellAlignmentInput(false, null, null, null, null), null, null, null, null);
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.ApplyStyle(tooLong, "A1:B2", style));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.SetComment(
                tooLong, "A1", new CommentInput("Review", "GridGrind", null)));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.ClearComment(tooLong, "A1"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.AppendRow(tooLong, List.of(new CellInput.Text("x"))));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.AutoSizeColumns(tooLong));
  }

  @Test
  void rejectsInvalidExcelSheetNameCharactersAndEdgeQuotes() {
    IllegalArgumentException invalidCharacter =
        assertThrows(
            IllegalArgumentException.class, () -> new WorkbookOperation.EnsureSheet("Bad:Name"));
    assertTrue(invalidCharacter.getMessage().contains("invalid Excel character ':'"));

    IllegalArgumentException leadingQuote =
        assertThrows(
            IllegalArgumentException.class, () -> new WorkbookOperation.EnsureSheet("'Budget"));
    assertTrue(leadingQuote.getMessage().contains("single quote"));

    IllegalArgumentException trailingQuote =
        assertThrows(
            IllegalArgumentException.class,
            () -> new WorkbookOperation.RenameSheet("Budget", "Budget'"));
    assertTrue(trailingQuote.getMessage().contains("single quote"));
  }

  @Test
  void operationTypeCoversAllSubtypes() {
    CellInput textValue = new CellInput.Text("x");
    CellStyleInput style =
        new CellStyleInput(
            null, new CellAlignmentInput(false, null, null, null, null), null, null, null, null);

    assertEquals("ENSURE_SHEET", new WorkbookOperation.EnsureSheet("Budget").operationType());
    assertEquals(
        "RENAME_SHEET", new WorkbookOperation.RenameSheet("Budget", "Summary").operationType());
    assertEquals("DELETE_SHEET", new WorkbookOperation.DeleteSheet("Budget").operationType());
    assertEquals("MOVE_SHEET", new WorkbookOperation.MoveSheet("Budget", 0).operationType());
    assertEquals(
        "COPY_SHEET",
        new WorkbookOperation.CopySheet(
                "Budget", "Budget Copy", new SheetCopyPosition.AppendAtEnd())
            .operationType());
    assertEquals(
        "SET_ACTIVE_SHEET", new WorkbookOperation.SetActiveSheet("Budget").operationType());
    assertEquals(
        "SET_SELECTED_SHEETS",
        new WorkbookOperation.SetSelectedSheets(List.of("Budget")).operationType());
    assertEquals(
        "SET_SHEET_VISIBILITY",
        new WorkbookOperation.SetSheetVisibility("Budget", ExcelSheetVisibility.HIDDEN)
            .operationType());
    assertEquals(
        "SET_SHEET_PROTECTION",
        new WorkbookOperation.SetSheetProtection("Budget", protectionSettings()).operationType());
    assertEquals(
        "CLEAR_SHEET_PROTECTION",
        new WorkbookOperation.ClearSheetProtection("Budget").operationType());
    assertEquals(
        "MERGE_CELLS", new WorkbookOperation.MergeCells("Budget", "A1:B2").operationType());
    assertEquals(
        "UNMERGE_CELLS", new WorkbookOperation.UnmergeCells("Budget", "A1:B2").operationType());
    assertEquals(
        "SET_COLUMN_WIDTH",
        new WorkbookOperation.SetColumnWidth("Budget", 0, 1, 16.0).operationType());
    assertEquals(
        "SET_ROW_HEIGHT", new WorkbookOperation.SetRowHeight("Budget", 0, 1, 28.5).operationType());
    assertEquals("INSERT_ROWS", new WorkbookOperation.InsertRows("Budget", 1, 2).operationType());
    assertEquals(
        "DELETE_ROWS",
        new WorkbookOperation.DeleteRows("Budget", new RowSpanInput.Band(1, 2)).operationType());
    assertEquals(
        "SHIFT_ROWS",
        new WorkbookOperation.ShiftRows("Budget", new RowSpanInput.Band(1, 2), 1).operationType());
    assertEquals(
        "INSERT_COLUMNS", new WorkbookOperation.InsertColumns("Budget", 1, 2).operationType());
    assertEquals(
        "DELETE_COLUMNS",
        new WorkbookOperation.DeleteColumns("Budget", new ColumnSpanInput.Band(1, 2))
            .operationType());
    assertEquals(
        "SHIFT_COLUMNS",
        new WorkbookOperation.ShiftColumns("Budget", new ColumnSpanInput.Band(1, 2), -1)
            .operationType());
    assertEquals(
        "SET_ROW_VISIBILITY",
        new WorkbookOperation.SetRowVisibility("Budget", new RowSpanInput.Band(1, 2), true)
            .operationType());
    assertEquals(
        "SET_COLUMN_VISIBILITY",
        new WorkbookOperation.SetColumnVisibility("Budget", new ColumnSpanInput.Band(1, 2), true)
            .operationType());
    assertEquals(
        "GROUP_ROWS",
        new WorkbookOperation.GroupRows("Budget", new RowSpanInput.Band(1, 2), false)
            .operationType());
    assertEquals(
        "UNGROUP_ROWS",
        new WorkbookOperation.UngroupRows("Budget", new RowSpanInput.Band(1, 2)).operationType());
    assertEquals(
        "GROUP_COLUMNS",
        new WorkbookOperation.GroupColumns("Budget", new ColumnSpanInput.Band(1, 2), false)
            .operationType());
    assertEquals(
        "UNGROUP_COLUMNS",
        new WorkbookOperation.UngroupColumns("Budget", new ColumnSpanInput.Band(1, 2))
            .operationType());
    assertEquals(
        "SET_SHEET_PANE",
        new WorkbookOperation.SetSheetPane("Budget", new PaneInput.Frozen(1, 2, 1, 2))
            .operationType());
    assertEquals(
        "SET_SHEET_ZOOM", new WorkbookOperation.SetSheetZoom("Budget", 125).operationType());
    assertEquals(
        "SET_PRINT_LAYOUT",
        new WorkbookOperation.SetPrintLayout("Budget", defaultPrintLayout()).operationType());
    assertEquals(
        "CLEAR_PRINT_LAYOUT", new WorkbookOperation.ClearPrintLayout("Budget").operationType());
    assertEquals(
        "SET_CELL", new WorkbookOperation.SetCell("Budget", "A1", textValue).operationType());
    assertEquals(
        "SET_RANGE",
        new WorkbookOperation.SetRange("Budget", "A1", List.of(List.of(textValue)))
            .operationType());
    assertEquals("CLEAR_RANGE", new WorkbookOperation.ClearRange("Budget", "A1").operationType());
    assertEquals(
        "SET_HYPERLINK",
        new WorkbookOperation.SetHyperlink(
                "Budget", "A1", new HyperlinkTarget.Url("https://example.com"))
            .operationType());
    assertEquals(
        "CLEAR_HYPERLINK", new WorkbookOperation.ClearHyperlink("Budget", "A1").operationType());
    assertEquals(
        "SET_COMMENT",
        new WorkbookOperation.SetComment(
                "Budget", "A1", new CommentInput("Review", "GridGrind", null))
            .operationType());
    assertEquals(
        "CLEAR_COMMENT", new WorkbookOperation.ClearComment("Budget", "A1").operationType());
    DrawingAnchorInput.TwoCell drawingAnchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(1, 2, 0, 0),
            new DrawingMarkerInput(2, 3, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    PictureDataInput pictureData =
        new PictureDataInput(
            ExcelPictureFormat.PNG,
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");
    assertEquals(
        "SET_PICTURE",
        new WorkbookOperation.SetPicture(
                "Budget",
                new PictureInput("OpsPicture", pictureData, drawingAnchor, "Queue preview"))
            .operationType());
    assertEquals(
        "SET_SHAPE",
        new WorkbookOperation.SetShape(
                "Budget",
                new ShapeInput(
                    "OpsShape",
                    ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                    drawingAnchor,
                    "rect",
                    "Queue"))
            .operationType());
    assertEquals(
        "SET_EMBEDDED_OBJECT",
        new WorkbookOperation.SetEmbeddedObject(
                "Budget",
                new EmbeddedObjectInput(
                    "OpsEmbed",
                    "Payload",
                    "payload.txt",
                    "payload.txt",
                    "cGF5bG9hZA==",
                    pictureData,
                    drawingAnchor))
            .operationType());
    assertEquals(
        "SET_DRAWING_OBJECT_ANCHOR",
        new WorkbookOperation.SetDrawingObjectAnchor("Budget", "OpsPicture", drawingAnchor)
            .operationType());
    assertEquals(
        "DELETE_DRAWING_OBJECT",
        new WorkbookOperation.DeleteDrawingObject("Budget", "OpsPicture").operationType());
    assertEquals(
        "APPLY_STYLE", new WorkbookOperation.ApplyStyle("Budget", "A1", style).operationType());
    assertEquals(
        "SET_DATA_VALIDATION",
        new WorkbookOperation.SetDataValidation("Budget", "A1", protocolValidation())
            .operationType());
    assertEquals(
        "CLEAR_DATA_VALIDATIONS",
        new WorkbookOperation.ClearDataValidations("Budget", new RangeSelection.All())
            .operationType());
    assertEquals(
        "SET_NAMED_RANGE",
        new WorkbookOperation.SetNamedRange(
                "BudgetTotal", new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "B4"))
            .operationType());
    assertEquals(
        "DELETE_NAMED_RANGE",
        new WorkbookOperation.DeleteNamedRange("BudgetTotal", new NamedRangeScope.Workbook())
            .operationType());
    assertEquals(
        "APPEND_ROW",
        new WorkbookOperation.AppendRow("Budget", List.of(textValue)).operationType());
    assertEquals(
        "AUTO_SIZE_COLUMNS", new WorkbookOperation.AutoSizeColumns("Budget").operationType());
    assertEquals("EVALUATE_FORMULAS", new WorkbookOperation.EvaluateFormulas().operationType());
    assertEquals(
        "FORCE_FORMULA_RECALCULATION_ON_OPEN",
        new WorkbookOperation.ForceFormulaRecalculationOnOpen().operationType());
  }

  private static List<CellInput> protocolRowValues() {
    return new ArrayList<>(List.of(new CellInput.Text("Item")));
  }

  private static List<List<CellInput>> protocolRows() {
    return new ArrayList<>(
        List.of(
            new ArrayList<>(List.of(new CellInput.Text("Item"), new CellInput.Numeric(12.0))),
            new ArrayList<>(List.of(new CellInput.Text("Tax"), new CellInput.Numeric(3.0)))));
  }

  private static CellStyleInput protocolStyle() {
    return new CellStyleInput(
        "#,##0.00",
        new CellAlignmentInput(
            true, ExcelHorizontalAlignment.RIGHT, ExcelVerticalAlignment.CENTER, null, null),
        new CellFontInput(true, null, null, null, null, null, null),
        null,
        null,
        null);
  }

  private static DataValidationInput protocolValidation() {
    return new DataValidationInput(
        new DataValidationRuleInput.TextLength(ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
        true,
        false,
        new DataValidationPromptInput("Reason", "Keep the reason concise.", true),
        new DataValidationErrorAlertInput(
            ExcelDataValidationErrorStyle.STOP, "Too long", "Use 20 characters or fewer.", true));
  }

  private static SheetProtectionSettings protectionSettings() {
    return new SheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }

  private static PrintLayoutInput defaultPrintLayout() {
    return new PrintLayoutInput(
        new PrintAreaInput.Range("A1:C20"),
        ExcelPrintOrientation.LANDSCAPE,
        new PrintScalingInput.Fit(1, 0),
        new PrintTitleRowsInput.Band(0, 0),
        new PrintTitleColumnsInput.Band(0, 0),
        new HeaderFooterTextInput("Budget", "", ""),
        new HeaderFooterTextInput("", "Page &P", ""));
  }
}
