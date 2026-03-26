package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Covers deterministic `.xlsx` round-trip cases that previously surfaced as fuzz findings. */
class XlsxRoundTripVerifierTest {
  /** Preserves alignment-only style patches through save and reopen. */
  @Test
  void requireRoundTripReadable_preservesAlignmentOnlyStyle(@TempDir Path tempDirectory)
      throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("Sheet1"),
            new WorkbookCommand.ApplyStyle(
                "Sheet1",
                "A1",
                ExcelCellStyle.alignment(
                    ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.TOP)));

    Path workbookPath = saveWorkbook(tempDirectory, commands);

    assertDoesNotThrow(() -> XlsxRoundTripVerifier.requireRoundTripReadable(workbookPath, commands));
  }

  /** Drops stale style expectations when a later value write resets the targeted cell. */
  @Test
  void requireRoundTripReadable_ignoresStylesClearedByLaterValueWrite(@TempDir Path tempDirectory)
      throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("Sheet1"),
            new WorkbookCommand.ApplyStyle(
                "Sheet1",
                "A1",
                new ExcelCellStyle(
                    null,
                    null,
                    null,
                    null,
                    ExcelHorizontalAlignment.CENTER,
                    ExcelVerticalAlignment.TOP,
                    null,
                    null,
                    null,
                    null,
                    null,
                    "#AABBCC",
                    null)),
            new WorkbookCommand.SetCell("Sheet1", "A1", ExcelCellValue.text("reset")));

    Path workbookPath = saveWorkbook(tempDirectory, commands);

    assertDoesNotThrow(() -> XlsxRoundTripVerifier.requireRoundTripReadable(workbookPath, commands));
  }

  /** Preserves richer Wave 2 formatting depth patches through save and reopen. */
  @Test
  void requireRoundTripReadable_preservesFormattingDepthStyle(@TempDir Path tempDirectory)
      throws IOException {
    ExcelCellStyle style =
        new ExcelCellStyle(
            null,
            Boolean.TRUE,
            Boolean.TRUE,
            Boolean.TRUE,
            ExcelHorizontalAlignment.CENTER,
            ExcelVerticalAlignment.TOP,
            "Arial",
            ExcelFontHeight.fromPoints(new BigDecimal("11.5")),
            "#112233",
            Boolean.TRUE,
            Boolean.TRUE,
            "#445566",
            new ExcelBorder(
                new ExcelBorderSide(ExcelBorderStyle.THIN),
                null,
                new ExcelBorderSide(ExcelBorderStyle.DOUBLE),
                null,
                null));
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("Sheet1"),
            new WorkbookCommand.ApplyStyle("Sheet1", "B2", style));

    Path workbookPath = saveWorkbook(tempDirectory, commands);

    assertDoesNotThrow(() -> XlsxRoundTripVerifier.requireRoundTripReadable(workbookPath, commands));
  }

  private static Path saveWorkbook(Path tempDirectory, List<WorkbookCommand> commands)
      throws IOException {
    Path workbookPath = tempDirectory.resolve("roundtrip.xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      new WorkbookCommandExecutor().apply(workbook, commands);
      workbook.save(workbookPath);
    }
    return workbookPath;
  }
}
