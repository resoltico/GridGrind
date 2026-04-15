package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.excel.ExcelCellFill;
import dev.erst.gridgrind.excel.ExcelCellFont;
import dev.erst.gridgrind.excel.ExcelCellProtection;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelGradientFill;
import dev.erst.gridgrind.excel.ExcelGradientStop;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelRichText;
import dev.erst.gridgrind.excel.ExcelRichTextRun;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Regression tests for the `.xlsx` round-trip verifier itself. */
class XlsxRoundTripVerifierTest {
  @Test
  void requireRoundTripReadableAcceptsGradientFills() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-gradient-roundtrip-",
                List.of(
                    new WorkbookCommand.CreateSheet("Budget"),
                    new WorkbookCommand.SetCell("Budget", "A1", ExcelCellValue.text("Gradient")),
                    new WorkbookCommand.ApplyStyle(
                        "Budget",
                        "A1",
                        new ExcelCellStyle(
                            null,
                            null,
                            null,
                            new ExcelCellFill(
                                null,
                                null,
                                null,
                                new ExcelGradientFill(
                                    "LINEAR",
                                    42.5d,
                                    null,
                                    null,
                                    null,
                                    null,
                                    List.of(
                                        new ExcelGradientStop(0.0d, new ExcelColor("#736C00")),
                                        new ExcelGradientStop(
                                            1.0d, new ExcelColor(null, 3, null, null))))),
                            null,
                            new ExcelCellProtection(true, true))))));
  }

  @Test
  void requireRoundTripReadableAcceptsMetadataAndPatternStyles() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-metadata-roundtrip-",
                List.of(
                    new WorkbookCommand.CreateSheet("Budget"),
                    new WorkbookCommand.SetCell(
                        "Budget",
                        "A1",
                        ExcelCellValue.richText(
                            new ExcelRichText(
                                List.of(
                                    new ExcelRichTextRun("Quarterly", null),
                                    new ExcelRichTextRun(
                                        " Report",
                                        new ExcelCellFont(
                                            true,
                                            null,
                                            null,
                                            null,
                                            new ExcelColor("#C0504D"),
                                            null,
                                            null)))))),
                    new WorkbookCommand.ApplyStyle(
                        "Budget",
                        "A1",
                        new ExcelCellStyle(
                            null,
                            null,
                            null,
                            new ExcelCellFill(ExcelFillPattern.BRICKS, "#102030", "#405060"),
                            null,
                            new ExcelCellProtection(false, true))),
                    new WorkbookCommand.SetHyperlink(
                        "Budget", "A1", new ExcelHyperlink.Url("https://example.com/report")),
                    new WorkbookCommand.SetComment(
                        "Budget", "A1", new ExcelComment("Review", "GridGrind", false)),
                    new WorkbookCommand.SetNamedRange(
                        new ExcelNamedRangeDefinition(
                            "BudgetTitle",
                            new ExcelNamedRangeScope.WorkbookScope(),
                            new ExcelNamedRangeTarget("Budget", "A1"))))));
  }

  @Test
  void requireRoundTripReadableAcceptsStructuredWorkbookFeatures() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-structured-roundtrip-",
                List.of(
                    new WorkbookCommand.CreateSheet("Budget"),
                    new WorkbookCommand.SetRange(
                        "Budget",
                        "A1:B3",
                        List.of(
                            List.of(ExcelCellValue.text("Item"), ExcelCellValue.text("Value")),
                            List.of(ExcelCellValue.text("Ada"), ExcelCellValue.number(49.0d)),
                            List.of(ExcelCellValue.text("Linus"), ExcelCellValue.number(10.0d)))),
                    new WorkbookCommand.SetCell("Budget", "C1", ExcelCellValue.text("scratch")),
                    new WorkbookCommand.ClearRange("Budget", "C1:C1"),
                    new WorkbookCommand.AppendRow(
                        "Budget",
                        List.of(ExcelCellValue.text("Grace"), ExcelCellValue.number(30.0d))),
                    new WorkbookCommand.SetDataValidation(
                        "Budget",
                        "B2:B4",
                        new ExcelDataValidationDefinition(
                            new ExcelDataValidationRule.TextLength(
                                ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
                            true,
                            false,
                            new ExcelDataValidationPrompt(
                                "Reason", "Use 20 characters or fewer.", true),
                            new ExcelDataValidationErrorAlert(
                                ExcelDataValidationErrorStyle.STOP,
                                "Too long",
                                "Use a shorter reason.",
                                true))),
                    new WorkbookCommand.SetConditionalFormatting(
                        "Budget",
                        new ExcelConditionalFormattingBlockDefinition(
                            List.of("B2:B4"),
                            List.of(
                                new ExcelConditionalFormattingRule.FormulaRule(
                                    "B2>0",
                                    true,
                                    new ExcelDifferentialStyle(
                                        "0.00", null, null, null, null, null, null, null, null))))),
                    new WorkbookCommand.SetAutofilter("Budget", "A1:B4"),
                    new WorkbookCommand.SetTable(
                        new ExcelTableDefinition(
                            "BudgetTable", "Budget", "A1:B4", true, new ExcelTableStyle.None())))));
  }

  private static void roundTrip(String prefix, List<WorkbookCommand> commands) throws IOException {
    Path workbookPath = Files.createTempFile(prefix, ".xlsx");
    Files.deleteIfExists(workbookPath);

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      new WorkbookCommandExecutor().apply(workbook, commands);
      workbook.save(workbookPath);
      XlsxRoundTripVerifier.requireRoundTripReadable(workbook, workbookPath, commands);
    }
  }
}
