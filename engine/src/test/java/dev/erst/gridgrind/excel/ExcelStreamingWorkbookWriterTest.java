package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;

/** Tests for low-memory append-only workbook writes backed by SXSSF. */
class ExcelStreamingWorkbookWriterTest {
  @Test
  void writesEverySupportedCellValueTypeAndMaterializesNestedOutputPaths() throws IOException {
    Path workbookPath =
        ExcelTempFiles.createManagedTempDirectory("gridgrind-streaming-write-")
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
      writer.markRecalculateOnOpen();
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
            || command instanceof WorkbookCommand.AppendRow) {
          continue;
        }

        IllegalArgumentException unsupported =
            assertThrows(
                IllegalArgumentException.class,
                () -> writer.apply(command),
                () ->
                    "expected STREAMING_WRITE rejection for " + command.getClass().getSimpleName());

        assertTrue(unsupported.getMessage().contains("execution.mode.writeMode=STREAMING_WRITE"));
        assertTrue(unsupported.getMessage().contains(command.commandType()));
      }
    } catch (IOException exception) {
      fail(exception);
    }
  }

  @Test
  void workbookCommandsExposeCanonicalOperationStyleNames() {
    assertEquals("ENSURE_SHEET", new WorkbookCommand.CreateSheet("Ops").commandType());
    assertEquals(
        "APPEND_ROW",
        new WorkbookCommand.AppendRow("Ops", List.of(ExcelCellValue.text("value"))).commandType());
    assertEquals(
        "SET_SHEET_PRESENTATION",
        new WorkbookCommand.SetSheetPresentation("Ops", ExcelSheetPresentation.defaults())
            .commandType());
    assertEquals(
        "SET_WORKBOOK_PROTECTION",
        new WorkbookCommand.SetWorkbookProtection(
                new ExcelWorkbookProtectionSettings(true, false, false, null, null))
            .commandType());
  }
}
