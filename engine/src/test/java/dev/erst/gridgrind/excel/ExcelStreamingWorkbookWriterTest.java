package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.lang.reflect.Method;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;

/** Tests for low-memory append-only workbook writes backed by SXSSF. */
@SuppressWarnings("PMD.AvoidAccessibilityAlteration")
class ExcelStreamingWorkbookWriterTest {
  @Test
  void writesEverySupportedCellValueTypeAndMaterializesNestedOutputPaths() throws IOException {
    Path workbookPath =
        Files.createTempDirectory("gridgrind-streaming-write-")
            .resolve("nested path")
            .resolve("streaming workbook.xlsx");

    try (ExcelStreamingWorkbookWriter writer = new ExcelStreamingWorkbookWriter()) {
      writer.apply(new WorkbookCommand.CreateSheet("Ops"));
      writer.apply(new WorkbookCommand.CreateSheet("Ops"));
      writer.apply(
          new WorkbookCommand.AppendRow(
              "Ops",
              List.of(
                  ExcelCellValue.blank(),
                  ExcelCellValue.text("Item"),
                  ExcelCellValue.richText(
                      new ExcelRichText(List.of(new ExcelRichTextRun("Rich", null)))),
                  ExcelCellValue.number(42.5d),
                  ExcelCellValue.bool(true),
                  ExcelCellValue.date(LocalDate.of(2026, 4, 13)),
                  ExcelCellValue.dateTime(LocalDateTime.of(2026, 4, 13, 9, 30, 15)),
                  ExcelCellValue.formula("2+3"))));
      writer.apply(
          new WorkbookCommand.AppendRow(
              "Ops", List.of(ExcelCellValue.text("Hosting"), ExcelCellValue.number(9.0d))));
      writer.apply(new WorkbookCommand.ForceFormulaRecalculationOnOpen());
      writer.save(workbookPath);
    }

    try (var workbook = WorkbookFactory.create(workbookPath.toFile())) {
      var sheet = workbook.getSheet("Ops");
      assertNotNull(sheet);
      assertEquals(CellType.BLANK, sheet.getRow(0).getCell(0).getCellType());
      assertEquals("Item", sheet.getRow(0).getCell(1).getStringCellValue());
      assertEquals("Rich", sheet.getRow(0).getCell(2).getRichStringCellValue().getString());
      assertEquals(42.5d, sheet.getRow(0).getCell(3).getNumericCellValue());
      assertTrue(sheet.getRow(0).getCell(4).getBooleanCellValue());
      assertEquals("yyyy-mm-dd", sheet.getRow(0).getCell(5).getCellStyle().getDataFormatString());
      assertEquals(
          "yyyy-mm-dd hh:mm:ss", sheet.getRow(0).getCell(6).getCellStyle().getDataFormatString());
      assertEquals("2+3", sheet.getRow(0).getCell(7).getCellFormula());
      assertEquals("Hosting", sheet.getRow(1).getCell(0).getStringCellValue());
      assertEquals(9.0d, sheet.getRow(1).getCell(1).getNumericCellValue());
      assertTrue(workbook.getForceFormulaRecalculation());
    }
  }

  @Test
  void rejectsMissingSheetAndEveryUnsupportedCommandVariant() {
    try (ExcelStreamingWorkbookWriter writer = new ExcelStreamingWorkbookWriter()) {
      SheetNotFoundException missingSheet =
          assertThrows(
              SheetNotFoundException.class,
              () ->
                  writer.apply(
                      new WorkbookCommand.AppendRow(
                          "Ops", List.of(ExcelCellValue.text("missing")))));
      assertEquals("Sheet does not exist: Ops", missingSheet.getMessage());

      for (WorkbookCommand command : WorkbookSampleFixtures.workbookCommands()) {
        if (command instanceof WorkbookCommand.CreateSheet
            || command instanceof WorkbookCommand.AppendRow
            || command instanceof WorkbookCommand.ForceFormulaRecalculationOnOpen) {
          continue;
        }

        IllegalArgumentException unsupported =
            assertThrows(
                IllegalArgumentException.class,
                () -> writer.apply(command),
                () ->
                    "expected STREAMING_WRITE rejection for " + command.getClass().getSimpleName());

        assertTrue(unsupported.getMessage().contains("executionMode.writeMode=STREAMING_WRITE"));
        assertTrue(unsupported.getMessage().contains(expectedCommandType(command)));
      }
    } catch (IOException exception) {
      fail(exception);
    }
  }

  @Test
  void commandTypeNamesSupportedStreamingCommands() throws Exception {
    assertEquals("ENSURE_SHEET", invokeCommandType(new WorkbookCommand.CreateSheet("Ops")));
    assertEquals(
        "APPEND_ROW",
        invokeCommandType(
            new WorkbookCommand.AppendRow("Ops", List.of(ExcelCellValue.text("value")))));
    assertEquals(
        "SET_SHEET_PRESENTATION",
        invokeCommandType(
            new WorkbookCommand.SetSheetPresentation("Ops", ExcelSheetPresentation.defaults())));
    assertEquals(
        "FORCE_FORMULA_RECALC_ON_OPEN",
        invokeCommandType(new WorkbookCommand.ForceFormulaRecalculationOnOpen()));
  }

  private static String expectedCommandType(WorkbookCommand command) {
    return switch (command) {
      case WorkbookCommand.CreateSheet _ -> "ENSURE_SHEET";
      case WorkbookCommand.AppendRow _ -> "APPEND_ROW";
      case WorkbookCommand.ForceFormulaRecalculationOnOpen _ -> "FORCE_FORMULA_RECALC_ON_OPEN";
      case WorkbookCommand.RenameSheet _ -> "RENAME_SHEET";
      case WorkbookCommand.DeleteSheet _ -> "DELETE_SHEET";
      case WorkbookCommand.MoveSheet _ -> "MOVE_SHEET";
      case WorkbookCommand.CopySheet _ -> "COPY_SHEET";
      case WorkbookCommand.SetActiveSheet _ -> "SET_ACTIVE_SHEET";
      case WorkbookCommand.SetSelectedSheets _ -> "SET_SELECTED_SHEETS";
      case WorkbookCommand.SetSheetVisibility _ -> "SET_SHEET_VISIBILITY";
      case WorkbookCommand.SetSheetProtection _ -> "SET_SHEET_PROTECTION";
      case WorkbookCommand.ClearSheetProtection _ -> "CLEAR_SHEET_PROTECTION";
      case WorkbookCommand.SetWorkbookProtection _ -> "SET_WORKBOOK_PROTECTION";
      case WorkbookCommand.ClearWorkbookProtection _ -> "CLEAR_WORKBOOK_PROTECTION";
      case WorkbookCommand.MergeCells _ -> "MERGE_CELLS";
      case WorkbookCommand.UnmergeCells _ -> "UNMERGE_CELLS";
      case WorkbookCommand.SetColumnWidth _ -> "SET_COLUMN_WIDTH";
      case WorkbookCommand.SetRowHeight _ -> "SET_ROW_HEIGHT";
      case WorkbookCommand.InsertRows _ -> "INSERT_ROWS";
      case WorkbookCommand.DeleteRows _ -> "DELETE_ROWS";
      case WorkbookCommand.ShiftRows _ -> "SHIFT_ROWS";
      case WorkbookCommand.InsertColumns _ -> "INSERT_COLUMNS";
      case WorkbookCommand.DeleteColumns _ -> "DELETE_COLUMNS";
      case WorkbookCommand.ShiftColumns _ -> "SHIFT_COLUMNS";
      case WorkbookCommand.SetRowVisibility _ -> "SET_ROW_VISIBILITY";
      case WorkbookCommand.SetColumnVisibility _ -> "SET_COLUMN_VISIBILITY";
      case WorkbookCommand.GroupRows _ -> "GROUP_ROWS";
      case WorkbookCommand.UngroupRows _ -> "UNGROUP_ROWS";
      case WorkbookCommand.GroupColumns _ -> "GROUP_COLUMNS";
      case WorkbookCommand.UngroupColumns _ -> "UNGROUP_COLUMNS";
      case WorkbookCommand.SetSheetPane _ -> "SET_SHEET_PANE";
      case WorkbookCommand.SetSheetZoom _ -> "SET_SHEET_ZOOM";
      case WorkbookCommand.SetSheetPresentation _ -> "SET_SHEET_PRESENTATION";
      case WorkbookCommand.SetPrintLayout _ -> "SET_PRINT_LAYOUT";
      case WorkbookCommand.ClearPrintLayout _ -> "CLEAR_PRINT_LAYOUT";
      case WorkbookCommand.SetCell _ -> "SET_CELL";
      case WorkbookCommand.SetRange _ -> "SET_RANGE";
      case WorkbookCommand.ClearRange _ -> "CLEAR_RANGE";
      case WorkbookCommand.SetHyperlink _ -> "SET_HYPERLINK";
      case WorkbookCommand.ClearHyperlink _ -> "CLEAR_HYPERLINK";
      case WorkbookCommand.SetComment _ -> "SET_COMMENT";
      case WorkbookCommand.ClearComment _ -> "CLEAR_COMMENT";
      case WorkbookCommand.SetPicture _ -> "SET_PICTURE";
      case WorkbookCommand.SetChart _ -> "SET_CHART";
      case WorkbookCommand.SetPivotTable _ -> "SET_PIVOT_TABLE";
      case WorkbookCommand.SetShape _ -> "SET_SHAPE";
      case WorkbookCommand.SetEmbeddedObject _ -> "SET_EMBEDDED_OBJECT";
      case WorkbookCommand.SetDrawingObjectAnchor _ -> "SET_DRAWING_OBJECT_ANCHOR";
      case WorkbookCommand.DeleteDrawingObject _ -> "DELETE_DRAWING_OBJECT";
      case WorkbookCommand.ApplyStyle _ -> "APPLY_STYLE";
      case WorkbookCommand.SetDataValidation _ -> "SET_DATA_VALIDATION";
      case WorkbookCommand.ClearDataValidations _ -> "CLEAR_DATA_VALIDATIONS";
      case WorkbookCommand.SetConditionalFormatting _ -> "SET_CONDITIONAL_FORMATTING";
      case WorkbookCommand.ClearConditionalFormatting _ -> "CLEAR_CONDITIONAL_FORMATTING";
      case WorkbookCommand.SetAutofilter _ -> "SET_AUTOFILTER";
      case WorkbookCommand.ClearAutofilter _ -> "CLEAR_AUTOFILTER";
      case WorkbookCommand.SetTable _ -> "SET_TABLE";
      case WorkbookCommand.DeleteTable _ -> "DELETE_TABLE";
      case WorkbookCommand.DeletePivotTable _ -> "DELETE_PIVOT_TABLE";
      case WorkbookCommand.SetNamedRange _ -> "SET_NAMED_RANGE";
      case WorkbookCommand.DeleteNamedRange _ -> "DELETE_NAMED_RANGE";
      case WorkbookCommand.AutoSizeColumns _ -> "AUTO_SIZE_COLUMNS";
      case WorkbookCommand.EvaluateAllFormulas _ -> "EVALUATE_FORMULAS";
      case WorkbookCommand.EvaluateFormulaCells _ -> "EVALUATE_FORMULA_CELLS";
      case WorkbookCommand.ClearFormulaCaches _ -> "CLEAR_FORMULA_CACHES";
    };
  }

  private static String invokeCommandType(WorkbookCommand command)
      throws ReflectiveOperationException {
    Method commandType =
        ExcelStreamingWorkbookWriter.class.getDeclaredMethod("commandType", WorkbookCommand.class);
    commandType.setAccessible(true);
    return (String) commandType.invoke(null, command);
  }
}
