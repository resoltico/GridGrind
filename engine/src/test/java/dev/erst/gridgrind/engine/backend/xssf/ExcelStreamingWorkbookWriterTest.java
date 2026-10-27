package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.style.ResolvedRichTextRun;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.engine.backend.streaming.StreamingWorkbookWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;
import java.util.zip.ZipFile;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for low-memory append-only workbook writes backed by SXSSF. */
class ExcelStreamingWorkbookWriterTest {
  @Test
  void writesEverySupportedCellValueTypeAndMaterializesNestedOutputPaths() throws IOException {
    Path workbookPath =
        ManagedTempFiles.createManagedTempDirectory("gridgrind-streaming-write-")
            .resolve("nested path")
            .resolve("streaming workbook.xlsx");

    try (StreamingWorkbookWriter writer = new StreamingWorkbookWriter()) {
      writer.apply(
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new WorkbookMutationAction.EnsureSheet()));
      writer.apply(
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new WorkbookMutationAction.EnsureSheet()));
      writer.apply(
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new ResolvedMutationAction.AppendRow(
                  List.of(
                      new CellValue.Blank(),
                      new CellValue.Text("Item"),
                      new ResolvedRichText(
                          List.of(new ResolvedRichTextRun("Rich", Optional.empty()))),
                      new CellValue.NumberValue(42.5d),
                      new CellValue.BooleanValue(true),
                      new CellValue.ErrorValue("#REF!"),
                      new CellValueSpec.Date(LocalDate.of(2026, 4, 13)),
                      new CellValueSpec.DateTime(LocalDateTime.of(2026, 4, 13, 9, 30, 15)),
                      new ResolvedFormula("2+3"),
                      new ResolvedRawFormula("LAMBDA(x,x+1)(A1)")))));
      writer.apply(
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new ResolvedMutationAction.AppendRow(
                  List.of(new CellValue.Text("Hosting"), new CellValue.NumberValue(9.0d)))));
      writer.markRecalculateOnOpen();
      writer.save(
          workbookPath,
          dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
    }

    try (var workbook = WorkbookFactory.create(workbookPath.toFile())) {
      var sheet = workbook.getSheet("Ops");
      assertNotNull(sheet);
      assertEquals(CellType.BLANK, sheet.getRow(0).getCell(0).getCellType());
      assertEquals("Item", sheet.getRow(0).getCell(1).getStringCellValue());
      assertEquals("Rich", sheet.getRow(0).getCell(2).getRichStringCellValue().getString());
      assertEquals(42.5d, sheet.getRow(0).getCell(3).getNumericCellValue());
      assertTrue(sheet.getRow(0).getCell(4).getBooleanCellValue());
      assertEquals(CellType.ERROR, sheet.getRow(0).getCell(5).getCellType());
      assertEquals(FormulaError.REF.getCode(), sheet.getRow(0).getCell(5).getErrorCellValue());
      assertEquals("yyyy-mm-dd", sheet.getRow(0).getCell(6).getCellStyle().getDataFormatString());
      assertEquals(
          "yyyy-mm-dd hh:mm:ss", sheet.getRow(0).getCell(7).getCellStyle().getDataFormatString());
      assertEquals("2+3", sheet.getRow(0).getCell(8).getCellFormula());
      assertEquals("LAMBDA(x,x+1)(A1)", sheet.getRow(0).getCell(9).getCellFormula());
      assertEquals("Hosting", sheet.getRow(1).getCell(0).getStringCellValue());
      assertEquals(9.0d, sheet.getRow(1).getCell(1).getNumericCellValue());
      assertTrue(workbook.getForceFormulaRecalculation());
    }
  }

  @Test
  void rejectsMissingSheetAndEveryUnsupportedCommandVariant() {
    try (StreamingWorkbookWriter writer = new StreamingWorkbookWriter()) {
      SheetNotFoundException missingSheet =
          assertThrows(
              SheetNotFoundException.class,
              () ->
                  writer.apply(
                      new ResolvedStep(
                          new SheetSelector.ByName(new SheetName("Ops")),
                          new ResolvedMutationAction.AppendRow(
                              List.of(new CellValue.Text("missing"))))));
      assertEquals("Sheet does not exist: Ops", missingSheet.getMessage());

      for (ResolvedStep command : WorkbookSampleFixtures.workbookCommands()) {
        if (command.action() instanceof WorkbookMutationAction.EnsureSheet
            || command.action() instanceof ResolvedMutationAction.AppendRow) {
          continue;
        }

        IllegalArgumentException unsupported =
            assertThrows(
                IllegalArgumentException.class,
                () -> writer.apply(command),
                () ->
                    "expected STREAMING_WRITE rejection for " + command.getClass().getSimpleName());

        assertTrue(unsupported.getMessage().contains("execution.mode.type=STREAMING_WRITE"));
        assertTrue(
            unsupported
                .getMessage()
                .contains(
                    dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames
                        .mutationActionTypeName(command.action().authoredType())));
      }
    } catch (IOException exception) {
      fail(exception);
    }
  }

  @Test
  void streamingWriterUsesSharedStringsToAvoidInlineStringPackageInflation() throws IOException {
    Path tempDirectory = ManagedTempFiles.createManagedTempDirectory("gridgrind-streaming-size-");
    Path gridGrindWorkbookPath = tempDirectory.resolve("gridgrind-streaming.xlsx");
    Path inlineWorkbookPath = tempDirectory.resolve("poi-inline-streaming.xlsx");

    List<ResolvedCellWrite> repeatedValues =
        List.of(
            new CellValue.Text("Quarterly revenue forecast"),
            new CellValue.Text("North"),
            new CellValue.Text("Approved"),
            new CellValue.Text("Quarterly revenue forecast"),
            new CellValue.Text("North"),
            new CellValue.Text("Approved"));

    try (StreamingWorkbookWriter writer = new StreamingWorkbookWriter()) {
      ResolvedStep repeatedAppendRow =
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new ResolvedMutationAction.AppendRow(repeatedValues));
      writer.apply(
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new WorkbookMutationAction.EnsureSheet()));
      for (int rowIndex = 0; rowIndex < 4_000; rowIndex++) {
        writer.apply(repeatedAppendRow);
      }
      writer.save(
          gridGrindWorkbookPath,
          dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
    }

    try (SXSSFWorkbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), 100, true, false)) {
      var sheet = workbook.createSheet("Ops");
      for (int rowIndex = 0; rowIndex < 4_000; rowIndex++) {
        var row = sheet.createRow(rowIndex);
        for (int columnIndex = 0; columnIndex < repeatedValues.size(); columnIndex++) {
          row.createCell(columnIndex)
              .setCellValue(((CellValue.Text) repeatedValues.get(columnIndex)).value());
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

  @Test
  void repeatedEquivalentStreamingSavesProduceIdenticalBytesAndCreateNewRejectsCollisions()
      throws IOException {
    Path workbookPath =
        ManagedTempFiles.createManagedTempFile("gridgrind-streaming-deterministic-", ".xlsx");
    try {
      byte[] firstSaveBytes = saveStreamingBudgetWorkbook(workbookPath);
      byte[] secondSaveBytes = saveStreamingBudgetWorkbook(workbookPath);

      assertArrayEquals(firstSaveBytes, secondSaveBytes);

      try (StreamingWorkbookWriter writer = new StreamingWorkbookWriter()) {
        writer.apply(
            new ResolvedStep(
                new SheetSelector.ByName(new SheetName("Ops")),
                new WorkbookMutationAction.EnsureSheet()));
        writer.apply(
            new ResolvedStep(
                new SheetSelector.ByName(new SheetName("Ops")),
                new ResolvedMutationAction.AppendRow(List.of(new CellValue.Text("collision")))));
        assertThrows(
            FileAlreadyExistsException.class,
            () ->
                writer.save(
                    workbookPath,
                    dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                        .CREATE_NEW));
      }
    } finally {
      Files.deleteIfExists(workbookPath);
    }
  }

  private static byte[] saveStreamingBudgetWorkbook(Path workbookPath) throws IOException {
    try (StreamingWorkbookWriter writer = new StreamingWorkbookWriter()) {
      writer.apply(
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new WorkbookMutationAction.EnsureSheet()));
      writer.apply(
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new ResolvedMutationAction.AppendRow(
                  List.of(new CellValue.Text("Hosting"), new CellValue.NumberValue(9.0d)))));
      writer.markRecalculateOnOpen();
      writer.save(
          workbookPath,
          dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
    }
    return Files.readAllBytes(workbookPath);
  }
}
