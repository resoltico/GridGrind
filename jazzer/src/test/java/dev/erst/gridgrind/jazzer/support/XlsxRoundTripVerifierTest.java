package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellGradientFillSpec;
import dev.erst.gridgrind.core.model.style.CellGradientStopSpec;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.style.ResolvedRichTextRun;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelComparisonFormulaSupport;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
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
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC";

  @Test
  void requireRoundTripReadableAcceptsGradientFills() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-gradient-roundtrip-",
                List.of(
                    new WorkbookSheetCommand.CreateSheet("Budget"),
                    new WorkbookCellCommand.SetCell("Budget", "A1", new CellValue.Text("Gradient")),
                    new WorkbookFormattingCommand.ApplyStyle(
                        "Budget",
                        "A1",
                        new CellStylePatch(
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.of(
                                CellFillSpec.gradient(
                                    CellGradientFillSpec.linear(
                                        Optional.of(42.5d),
                                        List.of(
                                            new CellGradientStopSpec(
                                                0.0d, ColorSpec.rgb(new RgbHex("#736C00"))),
                                            new CellGradientStopSpec(1.0d, ColorSpec.theme(3)))))),
                            Optional.empty(),
                            Optional.of(
                                new CellProtectionPatch(Optional.of(true), Optional.of(true))))))));
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
                        new ResolvedRichText(
                            List.of(
                                new ResolvedRichTextRun("Quarterly", Optional.empty()),
                                new ResolvedRichTextRun(
                                    " Report",
                                    Optional.of(
                                        new CellFontPatch(
                                            Optional.of(true),
                                            Optional.empty(),
                                            Optional.empty(),
                                            Optional.empty(),
                                            Optional.of(ColorSpec.rgb(new RgbHex("#C0504D"))),
                                            Optional.empty(),
                                            Optional.empty())))))),
                    new WorkbookFormattingCommand.ApplyStyle(
                        "Budget",
                        "A1",
                        new CellStylePatch(
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.of(
                                CellFillSpec.patternColors(
                                    ExcelFillPattern.BRICKS,
                                    ColorSpec.rgb(new RgbHex("#102030")),
                                    ColorSpec.rgb(new RgbHex("#405060")))),
                            Optional.empty(),
                            Optional.of(
                                new CellProtectionPatch(Optional.of(false), Optional.of(true))))),
                    new WorkbookAnnotationCommand.SetHyperlink(
                        "Budget", "A1", new HyperlinkTarget.Url("https://example.com/report")),
                    new WorkbookAnnotationCommand.SetComment(
                        "Budget", "A1", new ResolvedComment("Review", "GridGrind", false)),
                    new WorkbookMetadataCommand.SetNamedRange(
                        new StructuredMutationAction.SetNamedRange(
                            new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetTitle"),
                            new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                            NamedRangeTarget.range(
                                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                                new dev.erst.gridgrind.core.vocabulary.RangeRef("A1")))))));
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
                            List.of(new CellValue.Text("Item"), new CellValue.Text("Value")),
                            List.of(new CellValue.Text("Ada"), new CellValue.NumberValue(49.0d)),
                            List.of(
                                new CellValue.Text("Linus"), new CellValue.NumberValue(10.0d)))),
                    new WorkbookCellCommand.SetCell("Budget", "C1", new CellValue.Text("scratch")),
                    new WorkbookCellCommand.ClearRange("Budget", "C1:C1"),
                    new WorkbookCellCommand.AppendRow(
                        "Budget",
                        List.of(new CellValue.Text("Grace"), new CellValue.NumberValue(30.0d))),
                    new WorkbookFormattingCommand.SetDataValidation(
                        "Budget",
                        "B2:B4",
                        new ResolvedDataValidation(
                            new DataValidationRuleSpec.TextLength(
                                ExcelComparisonOperator.LESS_OR_EQUAL,
                                ExcelComparisonFormulaSupport.normalizeFormula("20", "formula1"),
                                ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                                    Optional.empty(), "formula2")),
                            true,
                            false,
                            Optional.of(
                                new ResolvedValidationPrompt(
                                    "Reason", "Use 20 characters or fewer.", true)),
                            Optional.of(
                                new ResolvedValidationErrorAlert(
                                    ExcelDataValidationErrorStyle.STOP,
                                    "Too long",
                                    "Use a shorter reason.",
                                    true)))),
                    new WorkbookFormattingCommand.SetConditionalFormatting(
                        "Budget",
                        List.of("B2:B4").stream()
                            .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                            .toList(),
                        new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
                            List.of(
                                new ConditionalFormattingRuleSpec.FormulaRule(
                                    "B2>0",
                                    true,
                                    Optional.of(
                                        new DifferentialStylePatch(
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
                        new ResolvedTable(
                            new TableName("BudgetTable"),
                            new SheetName("Budget"),
                            new RangeRef("A1:B4"),
                            true,
                            new TableStyleSpec.None())))));
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
                        "LL", "E2", new ResolvedComment("Note BudgetTotal", "GridGrind", true)),
                    new WorkbookAnnotationCommand.SetComment(
                        "LL", "A2", new ResolvedComment("Note Report_Value", "GridGrind", true)),
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
                        new ResolvedPicture(
                            "QueuePreview",
                            new BinaryContent(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
                            ExcelPictureFormat.PNG,
                            new DrawingAnchorSpec.TwoCell(
                                new DrawingMarkerSpec(1, 1, 0, 0),
                                new DrawingMarkerSpec(4, 6, 0, 0),
                                dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior
                                    .MOVE_AND_RESIZE),
                            Optional.of("Queue preview"))),
                    new WorkbookDrawingCommand.SetPicture(
                        "Queue",
                        new ResolvedPicture(
                            "QueuePreview2",
                            new BinaryContent(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
                            ExcelPictureFormat.PNG,
                            new DrawingAnchorSpec.TwoCell(
                                new DrawingMarkerSpec(6, 1, 0, 0),
                                new DrawingMarkerSpec(9, 6, 0, 0),
                                dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior
                                    .MOVE_AND_RESIZE),
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
                        new ResolvedEmbeddedObject(
                            "QueueEmbed",
                            "Payload",
                            "payload.txt",
                            "payload.txt",
                            new BinaryContent(
                                "payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
                            ExcelPictureFormat.PNG,
                            new BinaryContent(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
                            new DrawingAnchorSpec.TwoCell(
                                new DrawingMarkerSpec(1, 1, 0, 0),
                                new DrawingMarkerSpec(4, 6, 0, 0),
                                dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior
                                    .MOVE_AND_RESIZE))),
                    new WorkbookAnnotationCommand.SetComment(
                        "Queue", "B2", new ResolvedComment("Queue note", "GridGrind", true)),
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
                                new CellValue.Text("Month"),
                                new CellValue.Text("Plan"),
                                new CellValue.Text("Actual")),
                            List.of(
                                new CellValue.Text("Jan"),
                                new CellValue.NumberValue(10.0d),
                                new CellValue.NumberValue(12.0d)),
                            List.of(
                                new CellValue.Text("Feb"),
                                new CellValue.NumberValue(18.0d),
                                new CellValue.NumberValue(16.0d)),
                            List.of(
                                new CellValue.Text("Mar"),
                                new CellValue.NumberValue(15.0d),
                                new CellValue.NumberValue(21.0d)))),
                    new WorkbookCellCommand.SetArrayFormula(
                        "Approvals", "D2:D4", new ResolvedArrayFormula("B2:B4*C2:C4")),
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

  private static ResolvedSignatureLine signatureLineDefinition() {
    return new ResolvedSignatureLine(
        "BudgetSignature",
        new DrawingAnchorSpec.TwoCell(
            new DrawingMarkerSpec(1, 1, 0, 0),
            new DrawingMarkerSpec(4, 6, 0, 0),
            dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
        false,
        Optional.ofNullable("Review the budget before signing."),
        Optional.ofNullable("Ada Lovelace"),
        Optional.ofNullable("Finance"),
        Optional.ofNullable("ada@example.com"),
        Optional.empty(),
        Optional.ofNullable("invalid"),
        Optional.of(ExcelPictureFormat.PNG),
        Optional.of(new BinaryContent(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64))));
  }
}
