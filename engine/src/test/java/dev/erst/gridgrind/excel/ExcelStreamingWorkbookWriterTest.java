package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.zip.ZipFile;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
      writer.apply(new WorkbookSheetCommand.CreateSheet("Ops"));
      writer.apply(new WorkbookSheetCommand.CreateSheet("Ops"));
      writer.apply(
          new WorkbookCellCommand.AppendRow(
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
          new WorkbookCellCommand.AppendRow(
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
                      new WorkbookCellCommand.AppendRow(
                          "Ops", List.of(ExcelCellValue.text("missing")))));
      assertEquals("Sheet does not exist: Ops", missingSheet.getMessage());

      for (WorkbookCommand command : WorkbookSampleFixtures.workbookCommands()) {
        if (command instanceof WorkbookSheetCommand.CreateSheet
            || command instanceof WorkbookCellCommand.AppendRow) {
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
    assertEquals("ENSURE_SHEET", new WorkbookSheetCommand.CreateSheet("Ops").commandType());
    assertEquals(
        "APPEND_ROW",
        new WorkbookCellCommand.AppendRow("Ops", List.of(ExcelCellValue.text("value")))
            .commandType());
    assertEquals(
        "SET_SHEET_PRESENTATION",
        new WorkbookLayoutCommand.SetSheetPresentation("Ops", ExcelSheetPresentation.defaults())
            .commandType());
    assertEquals(
        "SET_WORKBOOK_PROTECTION",
        new WorkbookSheetCommand.SetWorkbookProtection(
                new ExcelWorkbookProtectionSettings(true, false, false, null, null))
            .commandType());
  }

  @Test
  void streamingWriterUsesSharedStringsToAvoidInlineStringPackageInflation() throws IOException {
    Path tempDirectory = ExcelTempFiles.createManagedTempDirectory("gridgrind-streaming-size-");
    Path gridGrindWorkbookPath = tempDirectory.resolve("gridgrind-streaming.xlsx");
    Path inlineWorkbookPath = tempDirectory.resolve("poi-inline-streaming.xlsx");

    List<ExcelCellValue> repeatedValues =
        List.of(
            ExcelCellValue.text("Quarterly revenue forecast"),
            ExcelCellValue.text("North"),
            ExcelCellValue.text("Approved"),
            ExcelCellValue.text("Quarterly revenue forecast"),
            ExcelCellValue.text("North"),
            ExcelCellValue.text("Approved"));

    try (ExcelStreamingWorkbookWriter writer = new ExcelStreamingWorkbookWriter()) {
      WorkbookCellCommand.AppendRow repeatedAppendRow =
          new WorkbookCellCommand.AppendRow("Ops", repeatedValues);
      writer.apply(new WorkbookSheetCommand.CreateSheet("Ops"));
      for (int rowIndex = 0; rowIndex < 4_000; rowIndex++) {
        writer.apply(repeatedAppendRow);
      }
      writer.save(gridGrindWorkbookPath);
    }

    try (SXSSFWorkbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), 100, true, false)) {
      var sheet = workbook.createSheet("Ops");
      for (int rowIndex = 0; rowIndex < 4_000; rowIndex++) {
        var row = sheet.createRow(rowIndex);
        for (int columnIndex = 0; columnIndex < repeatedValues.size(); columnIndex++) {
          row.createCell(columnIndex)
              .setCellValue(((ExcelCellValue.TextValue) repeatedValues.get(columnIndex)).value());
        }
      }
      try (OutputStream outputStream = java.nio.file.Files.newOutputStream(inlineWorkbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ZipFile archive = new ZipFile(gridGrindWorkbookPath.toFile())) {
      assertNotNull(archive.getEntry("xl/sharedStrings.xml"));
    }

    long gridGrindWorkbookSize = java.nio.file.Files.size(gridGrindWorkbookPath);
    long inlineWorkbookSize = java.nio.file.Files.size(inlineWorkbookPath);
    assertTrue(
        gridGrindWorkbookSize < inlineWorkbookSize,
        () ->
            "shared-strings streaming output should be smaller than inline-string SXSSF output: "
                + gridGrindWorkbookSize
                + " vs "
                + inlineWorkbookSize);
  }
}
