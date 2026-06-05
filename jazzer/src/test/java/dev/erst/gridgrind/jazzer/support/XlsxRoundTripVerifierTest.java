package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.excel.*;
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
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
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
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.drawing.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelSignatureLineDefinition;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationRule;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import java.util.Optional;
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
                    new WorkbookSheetCommand.CreateSheet("Budget"),
                    new WorkbookCellCommand.SetCell(
                        "Budget", "A1", ExcelCellValue.text("Gradient")),
                    new WorkbookFormattingCommand.ApplyStyle(
                        "Budget",
                        "A1",
                        new ExcelCellStyle(
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.of(
                                ExcelCellFill.gradient(
                                    ExcelGradientFill.linear(
                                        Optional.of(42.5d),
                                        List.of(
                                            new ExcelGradientStop(0.0d, ExcelColor.rgb("#736C00")),
                                            new ExcelGradientStop(1.0d, ExcelColor.theme(3)))))),
                            Optional.empty(),
                            Optional.of(
                                new ExcelCellProtection(Optional.of(true), Optional.of(true))))))));
  }

  @Test
  void requireRoundTripReadableAcceptsMetadataAndPatternStyles() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-metadata-roundtrip-",
                List.of(
                    new WorkbookSheetCommand.CreateSheet("Budget"),
                    new WorkbookCellCommand.SetCell(
                        "Budget",
                        "A1",
                        ExcelCellValue.richText(
                            new ExcelRichText(
                                List.of(
                                    new ExcelRichTextRun("Quarterly", Optional.empty()),
                                    new ExcelRichTextRun(
                                        " Report",
                                        Optional.of(
                                            new ExcelCellFont(
                                                Optional.of(true),
                                                Optional.empty(),
                                                Optional.empty(),
                                                Optional.empty(),
                                                Optional.of(ExcelColor.rgb("#C0504D")),
                                                Optional.empty(),
                                                Optional.empty()))))))),
                    new WorkbookFormattingCommand.ApplyStyle(
                        "Budget",
                        "A1",
                        new ExcelCellStyle(
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.of(
                                ExcelCellFill.patternColors(
                                    ExcelFillPattern.BRICKS,
                                    ExcelColor.rgb("#102030"),
                                    ExcelColor.rgb("#405060"))),
                            Optional.empty(),
                            Optional.of(
                                new ExcelCellProtection(Optional.of(false), Optional.of(true))))),
                    new WorkbookAnnotationCommand.SetHyperlink(
                        "Budget", "A1", new ExcelHyperlink.Url("https://example.com/report")),
                    new WorkbookAnnotationCommand.SetComment(
                        "Budget", "A1", new ExcelComment("Review", "GridGrind", false)),
                    new WorkbookMetadataCommand.SetNamedRange(
                        new ExcelNamedRangeDefinition(
                            "BudgetTitle",
                            new ExcelNamedRangeScope.WorkbookScope(),
                            ExcelNamedRangeTarget.range("Budget", "A1"))))));
  }

  @Test
  void requireRoundTripReadableAcceptsStructuredWorkbookFeatures() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-structured-roundtrip-",
                List.of(
                    new WorkbookSheetCommand.CreateSheet("Budget"),
                    new WorkbookCellCommand.SetRange(
                        "Budget",
                        "A1:B3",
                        List.of(
                            List.of(ExcelCellValue.text("Item"), ExcelCellValue.text("Value")),
                            List.of(ExcelCellValue.text("Ada"), ExcelCellValue.number(49.0d)),
                            List.of(ExcelCellValue.text("Linus"), ExcelCellValue.number(10.0d)))),
                    new WorkbookCellCommand.SetCell("Budget", "C1", ExcelCellValue.text("scratch")),
                    new WorkbookCellCommand.ClearRange("Budget", "C1:C1"),
                    new WorkbookCellCommand.AppendRow(
                        "Budget",
                        List.of(ExcelCellValue.text("Grace"), ExcelCellValue.number(30.0d))),
                    new WorkbookFormattingCommand.SetDataValidation(
                        "Budget",
                        "B2:B4",
                        new ExcelDataValidationDefinition(
                            new ExcelDataValidationRule.TextLength(
                                ExcelComparisonOperator.LESS_OR_EQUAL, "20", Optional.empty()),
                            true,
                            false,
                            Optional.of(
                                new ExcelDataValidationPrompt(
                                    "Reason", "Use 20 characters or fewer.", true)),
                            Optional.of(
                                new ExcelDataValidationErrorAlert(
                                    ExcelDataValidationErrorStyle.STOP,
                                    "Too long",
                                    "Use a shorter reason.",
                                    true)))),
                    new WorkbookFormattingCommand.SetConditionalFormatting(
                        "Budget",
                        new ExcelConditionalFormattingBlockDefinition(
                            List.of("B2:B4"),
                            List.of(
                                new ExcelConditionalFormattingRule.FormulaRule(
                                    "B2>0",
                                    true,
                                    Optional.of(
                                        new ExcelDifferentialStyle(
                                            Optional.of("0.00"),
                                            Optional.empty(),
                                            Optional.empty(),
                                            Optional.empty(),
                                            Optional.empty(),
                                            Optional.empty(),
                                            Optional.empty(),
                                            Optional.empty(),
                                            Optional.empty())))))),
                    new WorkbookTabularCommand.SetAutofilter("Budget", "A1:B4"),
                    new WorkbookTabularCommand.SetTable(
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
                    new WorkbookSheetCommand.CreateSheet("LL"),
                    new WorkbookAnnotationCommand.SetComment(
                        "LL", "E2", new ExcelComment("Note BudgetTotal", "GridGrind", true)),
                    new WorkbookAnnotationCommand.SetComment(
                        "LL", "A2", new ExcelComment("Note Report_Value", "GridGrind", true)),
                    new WorkbookSheetCommand.CreateSheet("LL"),
                    new WorkbookStructureCommand.DeleteColumns("LL", new ExcelColumnSpan(1, 3)),
                    new WorkbookStructureCommand.DeleteColumns("LL", new ExcelColumnSpan(0, 0)),
                    new WorkbookLayoutCommand.AutoSizeColumns("LL"),
                    new WorkbookLayoutCommand.AutoSizeColumns("LL"))));
  }

  @Test
  void requireRoundTripReadableAcceptsCopiedPicturesWithRetargetedRelations() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-copy-picture-roundtrip-",
                List.of(
                    new WorkbookSheetCommand.CreateSheet("Queue"),
                    new WorkbookDrawingCommand.SetPicture(
                        "Queue",
                        new ExcelPictureDefinition(
                            "QueuePreview",
                            new ExcelBinaryData(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
                            ExcelPictureFormat.PNG,
                            new ExcelDrawingAnchor.TwoCell(
                                new ExcelDrawingMarker(1, 1, 0, 0),
                                new ExcelDrawingMarker(4, 6, 0, 0),
                                null),
                            Optional.of("Queue preview"))),
                    new WorkbookDrawingCommand.SetPicture(
                        "Queue",
                        new ExcelPictureDefinition(
                            "QueuePreview2",
                            new ExcelBinaryData(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
                            ExcelPictureFormat.PNG,
                            new ExcelDrawingAnchor.TwoCell(
                                new ExcelDrawingMarker(6, 1, 0, 0),
                                new ExcelDrawingMarker(9, 6, 0, 0),
                                null),
                            Optional.empty())),
                    new WorkbookSheetCommand.CopySheet(
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
                    new WorkbookSheetCommand.CreateSheet("Queue"),
                    new WorkbookDrawingCommand.SetEmbeddedObject(
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
                    new WorkbookAnnotationCommand.SetComment(
                        "Queue", "B2", new ExcelComment("Queue note", "GridGrind", true)),
                    new WorkbookSheetCommand.CopySheet(
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
                    new WorkbookSheetCommand.CreateSheet("Approvals"),
                    new WorkbookCellCommand.SetRange(
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
                    new WorkbookCellCommand.SetArrayFormula(
                        "Approvals", "D2:D4", new ExcelArrayFormulaDefinition("B2:B4*C2:C4")),
                    new WorkbookCellCommand.ClearArrayFormula("Approvals", "D2"),
                    new WorkbookDrawingCommand.SetSignatureLine(
                        "Approvals", signatureLineDefinition()))));
  }

  private static void roundTrip(String prefix, List<WorkbookCommand> commands) throws IOException {
    Path workbookPath = Files.createTempFile(prefix, ".xlsx");
    Files.deleteIfExists(workbookPath);

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      new WorkbookExecutionEngine().apply(workbook, commands);
      JazzerWorkbookIoSupport.saveWorkbook(workbook, workbookPath);
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
        Optional.of(ExcelPictureFormat.PNG),
        Optional.of(new ExcelBinaryData(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64))));
  }
}
