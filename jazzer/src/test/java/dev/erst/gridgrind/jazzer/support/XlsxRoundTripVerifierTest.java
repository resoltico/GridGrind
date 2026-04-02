package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
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

  /** Accepts date-time append writes that relayer the required number format onto styled blank rows. */
  @Test
  void requireRoundTripReadable_acceptsDateTimeStyleRelayerOnStyledBlankAppendRow(
      @TempDir Path tempDirectory) throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("E"),
            new WorkbookCommand.AppendRow(
                "E",
                List.of(
                    ExcelCellValue.dateTime(java.time.LocalDateTime.of(2026, 2, 6, 13, 1, 1)),
                    ExcelCellValue.dateTime(java.time.LocalDateTime.of(2026, 2, 6, 13, 58, 1)))),
            new WorkbookCommand.ApplyStyle(
                "E",
                "A1:B2",
                new ExcelCellStyle(
                    "0.00",
                    Boolean.TRUE,
                    null,
                    Boolean.TRUE,
                    null,
                    ExcelVerticalAlignment.CENTER,
                    null,
                    null,
                    null,
                    null,
                    null,
                    "#603A79",
                    new ExcelBorder(
                        new ExcelBorderSide(ExcelBorderStyle.MEDIUM_DASHED),
                        null,
                        new ExcelBorderSide(ExcelBorderStyle.MEDIUM_DASHED),
                        null,
                        null))),
            new WorkbookCommand.AppendRow(
                "E",
                List.of(
                    ExcelCellValue.dateTime(java.time.LocalDateTime.of(2026, 2, 6, 13, 1, 58)),
                    ExcelCellValue.dateTime(java.time.LocalDateTime.of(2026, 2, 6, 13, 1, 1)))));

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

  /** Preserves normalized data-validation state through save and reopen. */
  @Test
  void requireRoundTripReadable_preservesDataValidationAuthoring(@TempDir Path tempDirectory)
      throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("Budget"),
            new WorkbookCommand.SetDataValidation(
                "Budget",
                "A2:C5",
                new ExcelDataValidationDefinition(
                    new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done")),
                    true,
                    false,
                    null,
                    null)),
            new WorkbookCommand.ClearDataValidations(
                "Budget", new ExcelRangeSelection.Selected(List.of("B3"))),
            new WorkbookCommand.SetDataValidation(
                "Budget",
                "E2:E5",
                new ExcelDataValidationDefinition(
                    new ExcelDataValidationRule.WholeNumber(
                        ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
                    false,
                    false,
                    null,
                    null)));

    Path workbookPath = saveWorkbook(tempDirectory, commands);

    assertDoesNotThrow(() -> XlsxRoundTripVerifier.requireRoundTripReadable(workbookPath, commands));
  }

  /** Preserves sheet autofilters and workbook tables through save and reopen. */
  @Test
  void requireRoundTripReadable_preservesConditionalFormattingAuthoring(@TempDir Path tempDirectory)
      throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("Budget"),
            new WorkbookCommand.SetRange(
                "Budget",
                "A1:B5",
                List.of(
                    List.of(ExcelCellValue.text("Status"), ExcelCellValue.text("Amount")),
                    List.of(ExcelCellValue.text("Queued"), ExcelCellValue.number(1.0)),
                    List.of(ExcelCellValue.text("Done"), ExcelCellValue.number(9.0)),
                    List.of(ExcelCellValue.text("Done"), ExcelCellValue.number(11.0)),
                    List.of(ExcelCellValue.text("Queued"), ExcelCellValue.number(4.0)))),
            new WorkbookCommand.SetConditionalFormatting(
                "Budget",
                new ExcelConditionalFormattingBlockDefinition(
                    List.of("A2:A5"),
                    List.of(
                        new ExcelConditionalFormattingRule.FormulaRule(
                            "A2=\"Done\"",
                            true,
                            new ExcelDifferentialStyle(
                                null, true, null, null, "#102030", null, null, "#E0F0AA", null))))),
            new WorkbookCommand.ClearConditionalFormatting(
                "Budget", new ExcelRangeSelection.Selected(List.of("A2:A5"))),
            new WorkbookCommand.SetConditionalFormatting(
                "Budget",
                new ExcelConditionalFormattingBlockDefinition(
                    List.of("B2:B5"),
                    List.of(
                        new ExcelConditionalFormattingRule.FormulaRule(
                            "B2>5",
                            true,
                            new ExcelDifferentialStyle(
                                "0.00", true, null, null, "#102030", null, null, "#E0F0AA", null)),
                        new ExcelConditionalFormattingRule.CellValueRule(
                            ExcelComparisonOperator.BETWEEN,
                            "1",
                            "10",
                            false,
                            new ExcelDifferentialStyle(
                                null, null, true, null, null, null, null, null, null))))));

    Path workbookPath = saveWorkbook(tempDirectory, commands);

    assertDoesNotThrow(() -> XlsxRoundTripVerifier.requireRoundTripReadable(workbookPath, commands));
  }

  /** Preserves sheet autofilters and workbook tables through save and reopen. */
  @Test
  void requireRoundTripReadable_preservesAutofilterAndTableAuthoring(
      @TempDir Path tempDirectory) throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("Budget"),
            new WorkbookCommand.SetRange(
                "Budget",
                "A1:C4",
                List.of(
                    List.of(
                        ExcelCellValue.text("Item"),
                        ExcelCellValue.text("Amount"),
                        ExcelCellValue.text("Billable")),
                    List.of(
                        ExcelCellValue.text("Hosting"),
                        ExcelCellValue.number(49.0),
                        ExcelCellValue.bool(true)),
                    List.of(
                        ExcelCellValue.text("Domain"),
                        ExcelCellValue.number(12.0),
                        ExcelCellValue.bool(false)),
                    List.of(
                        ExcelCellValue.text("Support"),
                        ExcelCellValue.number(18.0),
                        ExcelCellValue.bool(true)))),
            new WorkbookCommand.SetRange(
                "Budget",
                "E1:F3",
                List.of(
                    List.of(ExcelCellValue.text("Queue"), ExcelCellValue.text("Owner")),
                    List.of(
                        ExcelCellValue.text("Late invoices"), ExcelCellValue.text("Marta")),
                    List.of(
                        ExcelCellValue.text("Badge orders"), ExcelCellValue.text("Rihards")))),
            new WorkbookCommand.SetAutofilter("Budget", "E1:F3"),
            new WorkbookCommand.SetTable(
                new ExcelTableDefinition(
                    "BudgetTable",
                    "Budget",
                    "A1:C4",
                    false,
                    new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false))));

    Path workbookPath = saveWorkbook(tempDirectory, commands);

    assertDoesNotThrow(() -> XlsxRoundTripVerifier.requireRoundTripReadable(workbookPath, commands));
  }

  /** Accepts header rewrites after table creation without losing persisted table column names. */
  @Test
  void requireRoundTripReadable_preservesTableHeaderRewritesAfterTableCreation(
      @TempDir Path tempDirectory) throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("V"),
            new WorkbookCommand.SetRange(
                "V",
                "A1:B3",
                List.of(
                    List.of(ExcelCellValue.text("c]cc"), ExcelCellValue.text("Task")),
                    List.of(ExcelCellValue.text("Ada"), ExcelCellValue.text("Queue")),
                    List.of(ExcelCellValue.text("Lin"), ExcelCellValue.text("Pack")))),
            new WorkbookCommand.SetTable(
                new ExcelTableDefinition(
                    "BudgetTable",
                    "V",
                    "A1:B3",
                    false,
                    new ExcelTableStyle.Named("TableStyleMedium2", true, true, true, true))),
            new WorkbookCommand.SetRange(
                "V",
                "A1:B2",
                List.of(
                    List.of(ExcelCellValue.text("QQQQq"), ExcelCellValue.text("Task")),
                    List.of(ExcelCellValue.text("Ada"), ExcelCellValue.text("Queue")))));

    Path workbookPath = saveWorkbook(tempDirectory, commands);

    assertDoesNotThrow(() -> XlsxRoundTripVerifier.requireRoundTripReadable(workbookPath, commands));
  }

  /** Accepts header style patches that change the displayed table header text for typed cells. */
  @Test
  void requireRoundTripReadable_preservesTableHeadersAfterHeaderStyleDisplayChanges(
      @TempDir Path tempDirectory) throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("V"),
            new WorkbookCommand.AppendRow(
                "V",
                List.of(
                    ExcelCellValue.dateTime(LocalDateTime.of(2026, 2, 6, 0, 19, 17)),
                    ExcelCellValue.date(LocalDate.of(2026, 6, 26)))),
            new WorkbookCommand.SetCell("V", "A2", ExcelCellValue.text("Ada")),
            new WorkbookCommand.SetCell("V", "B2", ExcelCellValue.text("Queue")),
            new WorkbookCommand.SetCell("V", "A3", ExcelCellValue.text("Totals")),
            new WorkbookCommand.SetCell("V", "B3", ExcelCellValue.text("Done")),
            new WorkbookCommand.SetTable(
                new ExcelTableDefinition(
                    "OpsTable",
                    "V",
                    "A1:B3",
                    true,
                    new ExcelTableStyle.Named("TableStyleMedium2", true, true, true, true))),
            new WorkbookCommand.ApplyStyle("V", "A1:B2", ExcelCellStyle.numberFormat("yyyy-mm-dd")));

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

  /** Preserves copied-sheet order, active selection state, visibility, and protection through reopen. */
  @Test
  void requireRoundTripReadable_preservesB1SheetState(@TempDir Path tempDirectory)
      throws IOException {
    List<WorkbookCommand> commands =
        List.of(
            new WorkbookCommand.CreateSheet("Alpha"),
            new WorkbookCommand.CreateSheet("Beta"),
            new WorkbookCommand.SetCell("Alpha", "A1", ExcelCellValue.text("Queue")),
            new WorkbookCommand.SetCell("Alpha", "B2", ExcelCellValue.number(7.0)),
            new WorkbookCommand.CopySheet(
                "Alpha", "Replica", new ExcelSheetCopyPosition.AtIndex(1)),
            new WorkbookCommand.SetActiveSheet("Replica"),
            new WorkbookCommand.SetSelectedSheets(List.of("Alpha", "Replica")),
            new WorkbookCommand.SetSheetVisibility("Beta", ExcelSheetVisibility.VERY_HIDDEN),
            new WorkbookCommand.SetSheetProtection("Replica", protectionSettings()));

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

  private static ExcelSheetProtectionSettings protectionSettings() {
    return new ExcelSheetProtectionSettings(
        false,
        true,
        false,
        true,
        false,
        true,
        false,
        true,
        false,
        true,
        false,
        true,
        false,
        true,
        false);
  }
}
