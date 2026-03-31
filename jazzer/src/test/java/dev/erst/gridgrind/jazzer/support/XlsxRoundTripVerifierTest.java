package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
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

  /** Preserves style expectations when a later value write targets an already styled cell. */
  @Test
  void requireRoundTripReadable_preservesStylesAcrossLaterValueWrite(@TempDir Path tempDirectory)
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

  /** Preserves style state when APPEND_ROW writes into a style-only row selected by append semantics. */
  @Test
  void requireRoundTripReadable_preservesStylesWhenAppendRowReusesStyledBlankRow(
      @TempDir Path tempDirectory) throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("X"),
            new WorkbookCommand.ApplyStyle(
                "X",
                "A1:B2",
                new ExcelCellStyle(
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    "Aptos",
                    ExcelFontHeight.fromPoints(new BigDecimal("15.2")),
                    "#A3A3A3",
                    null,
                    Boolean.TRUE,
                    "#CDCDCD",
                    new ExcelBorder(
                        new ExcelBorderSide(ExcelBorderStyle.DASH_DOT), null, null, null, null))),
            new WorkbookCommand.AppendRow(
                "X",
                List.of(ExcelCellValue.number(607.8483822864587), ExcelCellValue.bool(false))));

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

  /** Preserves hyperlink, comment, and named-range state through save and reopen. */
  @Test
  void requireRoundTripReadable_preservesExcelAuthoringMetadata(@TempDir Path tempDirectory)
      throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("Sheet1"),
            new WorkbookCommand.SetCell("Sheet1", "A1", ExcelCellValue.text("Report")),
            new WorkbookCommand.SetHyperlink(
                "Sheet1", "A1", new ExcelHyperlink.Url("https://example.com/report")),
            new WorkbookCommand.SetComment(
                "Sheet1", "A1", new ExcelComment("Review", "GridGrind", true)),
            new WorkbookCommand.SetNamedRange(
                new ExcelNamedRangeDefinition(
                    "BudgetTotal",
                    new ExcelNamedRangeScope.WorkbookScope(),
                    new ExcelNamedRangeTarget("Sheet1", "A1"))),
            new WorkbookCommand.SetNamedRange(
                new ExcelNamedRangeDefinition(
                    "BudgetTable",
                    new ExcelNamedRangeScope.SheetScope("Sheet1"),
                    new ExcelNamedRangeTarget("Sheet1", "A1:B2"))));

    Path workbookPath = saveWorkbook(tempDirectory, commands);

    assertDoesNotThrow(() -> XlsxRoundTripVerifier.requireRoundTripReadable(workbookPath, commands));
  }

  /** Preserves the last hyperlink written to a cell through save and reopen. */
  @Test
  void requireRoundTripReadable_preservesLatestRepeatedHyperlinkWrite(@TempDir Path tempDirectory)
      throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("C"),
            new WorkbookCommand.SetHyperlink(
                "C", "F18", new ExcelHyperlink.Email("Report_Value@example.com")),
            new WorkbookCommand.SetHyperlink(
                "C", "F18", new ExcelHyperlink.Email("Summary.Total@example.com")));

    Path workbookPath = saveWorkbook(tempDirectory, commands);

    assertDoesNotThrow(() -> XlsxRoundTripVerifier.requireRoundTripReadable(workbookPath, commands));
  }

  /** Accepts named-range targets that are normalized during persistence and reopen. */
  @Test
  void requireRoundTripReadable_normalizesReversedNamedRangeTargets(@TempDir Path tempDirectory)
      throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("Sheet1"),
            new WorkbookCommand.SetNamedRange(
                new ExcelNamedRangeDefinition(
                    "BudgetTotal",
                    new ExcelNamedRangeScope.WorkbookScope(),
                    new ExcelNamedRangeTarget("Sheet1", "B2:A1"))));

    Path workbookPath = saveWorkbook(tempDirectory, commands);

    assertDoesNotThrow(() -> XlsxRoundTripVerifier.requireRoundTripReadable(workbookPath, commands));
  }

  /** Tracks renamed sheets and later metadata removals without keeping stale expectations. */
  @Test
  void requireRoundTripReadable_tracksRenamesAndMetadataDeletes(@TempDir Path tempDirectory)
      throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("Sheet1"),
            new WorkbookCommand.SetCell("Sheet1", "A1", ExcelCellValue.text("Report")),
            new WorkbookCommand.SetHyperlink(
                "Sheet1", "A1", new ExcelHyperlink.Document("Sheet1!A1")),
            new WorkbookCommand.SetComment(
                "Sheet1", "A1", new ExcelComment("Review", "GridGrind", false)),
            new WorkbookCommand.SetNamedRange(
                new ExcelNamedRangeDefinition(
                    "BudgetTotal",
                    new ExcelNamedRangeScope.WorkbookScope(),
                    new ExcelNamedRangeTarget("Sheet1", "A1"))),
            new WorkbookCommand.RenameSheet("Sheet1", "Summary"),
            new WorkbookCommand.ClearHyperlink("Summary", "A1"),
            new WorkbookCommand.ClearComment("Summary", "A1"),
            new WorkbookCommand.DeleteNamedRange(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope()));

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
