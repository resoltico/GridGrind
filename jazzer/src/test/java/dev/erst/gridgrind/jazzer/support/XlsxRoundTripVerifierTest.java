package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.excel.ExcelArrayFormulaDefinition;
import dev.erst.gridgrind.excel.ExcelBinaryData;
import dev.erst.gridgrind.excel.ExcelCellFill;
import dev.erst.gridgrind.excel.ExcelCellFont;
import dev.erst.gridgrind.excel.ExcelCellProtection;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.ExcelGradientFill;
import dev.erst.gridgrind.excel.ExcelGradientStop;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.ExcelRichText;
import dev.erst.gridgrind.excel.ExcelRichTextRun;
import dev.erst.gridgrind.excel.ExcelSignatureLineDefinition;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Regression tests for the `.xlsx` round-trip verifier itself. */
class XlsxRoundTripVerifierTest {
  private static final String PNG_PIXEL_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=";

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
                            ExcelCellFill.gradient(
                                ExcelGradientFill.linear(
                                    42.5d,
                                    List.of(
                                        new ExcelGradientStop(0.0d, ExcelColor.rgb("#736C00")),
                                        new ExcelGradientStop(1.0d, ExcelColor.theme(3))))),
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
                                            ExcelColor.rgb("#C0504D"),
                                            null,
                                            null)))))),
                    new WorkbookCommand.ApplyStyle(
                        "Budget",
                        "A1",
                        new ExcelCellStyle(
                            null,
                            null,
                            null,
                            ExcelCellFill.patternColors(
                                ExcelFillPattern.BRICKS,
                                ExcelColor.rgb("#102030"),
                                ExcelColor.rgb("#405060")),
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

  @Test
  void requireRoundTripReadableAcceptsCommentCollisionsDuringColumnEdits() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-comment-column-roundtrip-",
                List.of(
                    new WorkbookCommand.CreateSheet("LL"),
                    new WorkbookCommand.SetComment(
                        "LL", "E2", new ExcelComment("Note BudgetTotal", "GridGrind", true)),
                    new WorkbookCommand.SetComment(
                        "LL", "A2", new ExcelComment("Note Report_Value", "GridGrind", true)),
                    new WorkbookCommand.CreateSheet("LL"),
                    new WorkbookCommand.DeleteColumns("LL", new ExcelColumnSpan(1, 3)),
                    new WorkbookCommand.DeleteColumns("LL", new ExcelColumnSpan(0, 0)),
                    new WorkbookCommand.AutoSizeColumns("LL"),
                    new WorkbookCommand.AutoSizeColumns("LL"))));
  }

  @Test
  void requireRoundTripReadableAcceptsCopiedPicturesWithRetargetedRelations() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-copy-picture-roundtrip-",
                List.of(
                    new WorkbookCommand.CreateSheet("Queue"),
                    new WorkbookCommand.SetPicture(
                        "Queue",
                        new ExcelPictureDefinition(
                            "QueuePreview",
                            new ExcelBinaryData(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
                            ExcelPictureFormat.PNG,
                            new ExcelDrawingAnchor.TwoCell(
                                new ExcelDrawingMarker(1, 1, 0, 0),
                                new ExcelDrawingMarker(4, 6, 0, 0),
                                null),
                            "Queue preview")),
                    new WorkbookCommand.SetPicture(
                        "Queue",
                        new ExcelPictureDefinition(
                            "QueuePreview2",
                            new ExcelBinaryData(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
                            ExcelPictureFormat.PNG,
                            new ExcelDrawingAnchor.TwoCell(
                                new ExcelDrawingMarker(6, 1, 0, 0),
                                new ExcelDrawingMarker(9, 6, 0, 0),
                                null),
                            null)),
                    new WorkbookCommand.CopySheet(
                        "Queue",
                        "Queue Copy",
                        new dev.erst.gridgrind.excel.ExcelSheetCopyPosition.AppendAtEnd()))));
  }

  @Test
  void requireRoundTripReadableAcceptsCopiedEmbeddedObjectPreviewRelations() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-copy-embedded-preview-roundtrip-",
                List.of(
                    new WorkbookCommand.CreateSheet("Queue"),
                    new WorkbookCommand.SetEmbeddedObject(
                        "Queue",
                        new ExcelEmbeddedObjectDefinition(
                            "QueueEmbed",
                            "Payload",
                            "payload.txt",
                            "payload.txt",
                            new ExcelBinaryData(
                                "payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
                            ExcelPictureFormat.PNG,
                            new ExcelBinaryData(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
                            new ExcelDrawingAnchor.TwoCell(
                                new ExcelDrawingMarker(1, 1, 0, 0),
                                new ExcelDrawingMarker(4, 6, 0, 0),
                                null))),
                    new WorkbookCommand.SetComment(
                        "Queue", "B2", new ExcelComment("Queue note", "GridGrind", true)),
                    new WorkbookCommand.CopySheet(
                        "Queue",
                        "Queue Copy",
                        new dev.erst.gridgrind.excel.ExcelSheetCopyPosition.AppendAtEnd()))));
  }

  @Test
  void requireRoundTripReadableAcceptsArrayFormulasAndSignatureLines() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-array-signature-roundtrip-",
                List.of(
                    new WorkbookCommand.CreateSheet("Approvals"),
                    new WorkbookCommand.SetRange(
                        "Approvals",
                        "A1:C4",
                        List.of(
                            List.of(
                                ExcelCellValue.text("Month"),
                                ExcelCellValue.text("Plan"),
                                ExcelCellValue.text("Actual")),
                            List.of(
                                ExcelCellValue.text("Jan"),
                                ExcelCellValue.number(10.0d),
                                ExcelCellValue.number(12.0d)),
                            List.of(
                                ExcelCellValue.text("Feb"),
                                ExcelCellValue.number(18.0d),
                                ExcelCellValue.number(16.0d)),
                            List.of(
                                ExcelCellValue.text("Mar"),
                                ExcelCellValue.number(15.0d),
                                ExcelCellValue.number(21.0d)))),
                    new WorkbookCommand.SetArrayFormula(
                        "Approvals", "D2:D4", new ExcelArrayFormulaDefinition("B2:B4*C2:C4")),
                    new WorkbookCommand.ClearArrayFormula("Approvals", "D2"),
                    new WorkbookCommand.SetSignatureLine("Approvals", signatureLineDefinition()))));
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

  private static ExcelSignatureLineDefinition signatureLineDefinition() {
    return new ExcelSignatureLineDefinition(
        "BudgetSignature",
        new ExcelDrawingAnchor.TwoCell(
            new ExcelDrawingMarker(1, 1, 0, 0), new ExcelDrawingMarker(4, 6, 0, 0), null),
        false,
        "Review the budget before signing.",
        "Ada Lovelace",
        "Finance",
        "ada@example.com",
        null,
        "invalid",
        ExcelPictureFormat.PNG,
        new ExcelBinaryData(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64)));
  }
}
