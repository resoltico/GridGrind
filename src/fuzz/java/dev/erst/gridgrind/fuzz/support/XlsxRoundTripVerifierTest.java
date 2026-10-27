package dev.erst.gridgrind.fuzz.support;

import static dev.erst.gridgrind.fuzz.support.OperationSequenceCommandFactory.setNamedRangeStep;
import static dev.erst.gridgrind.fuzz.support.OperationSequenceCommandFactory.setTableStep;
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
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.backend.xssf.*;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import dev.erst.gridgrind.engine.formulas.ComparisonFormulas;
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
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("Budget")),
                        new WorkbookMutationAction.EnsureSheet()),
                    new ResolvedStep(
                        new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                        new ResolvedMutationAction.SetCell(new CellValue.Text("Gradient"))),
                    new ResolvedStep(
                        new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1")),
                        new CellMutationAction.ApplyStyle(
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
                                                new CellGradientStopSpec(
                                                    1.0d, ColorSpec.theme(3)))))),
                                Optional.empty(),
                                Optional.of(
                                    new CellProtectionPatch(
                                        Optional.of(true), Optional.of(true)))))))));
  }

  @Test
  void requireRoundTripReadableAcceptsMetadataAndPatternStyles() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-metadata-roundtrip-",
                List.of(
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("Budget")),
                        new WorkbookMutationAction.EnsureSheet()),
                    new ResolvedStep(
                        new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                        new ResolvedMutationAction.SetCell(
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
                                                Optional.empty()))))))),
                    new ResolvedStep(
                        new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1")),
                        new CellMutationAction.ApplyStyle(
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
                                    new CellProtectionPatch(
                                        Optional.of(false), Optional.of(true)))))),
                    new ResolvedStep(
                        new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                        new CellMutationAction.SetHyperlink(
                            new HyperlinkTarget.Url("https://example.com/report"))),
                    new ResolvedStep(
                        new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                        new ResolvedMutationAction.SetComment(
                            new ResolvedComment("Review", "GridGrind", false))),
                    setNamedRangeStep(
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
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("Budget")),
                        new WorkbookMutationAction.EnsureSheet()),
                    new ResolvedStep(
                        new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B3")),
                        new ResolvedMutationAction.SetRange(
                            List.of(
                                List.of(new CellValue.Text("Item"), new CellValue.Text("Value")),
                                List.of(
                                    new CellValue.Text("Ada"), new CellValue.NumberValue(49.0d)),
                                List.of(
                                    new CellValue.Text("Linus"),
                                    new CellValue.NumberValue(10.0d))))),
                    new ResolvedStep(
                        new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("C1")),
                        new ResolvedMutationAction.SetCell(new CellValue.Text("scratch"))),
                    new ResolvedStep(
                        new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("C1:C1")),
                        new CellMutationAction.ClearRange()),
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("Budget")),
                        new ResolvedMutationAction.AppendRow(
                            List.of(
                                new CellValue.Text("Grace"), new CellValue.NumberValue(30.0d)))),
                    new ResolvedStep(
                        new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("B2:B4")),
                        new ResolvedMutationAction.SetDataValidation(
                            new ResolvedDataValidation(
                                new DataValidationRuleSpec.TextLength(
                                    ExcelComparisonOperator.LESS_OR_EQUAL,
                                    ComparisonFormulas.normalizeFormula("20", "formula1"),
                                    ComparisonFormulas.normalizeOptionalFormula(
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
                                        true))))),
                    new ResolvedStep(
                        new RangeSelector.ByRanges(
                            new SheetName("Budget"),
                            List.of("B2:B4").stream()
                                .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                                .toList()),
                        new StructuredMutationAction.SetConditionalFormatting(
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
                                                Optional.empty()))))))),
                    new ResolvedStep(
                        new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B4")),
                        new StructuredMutationAction.SetAutofilter(List.of(), Optional.empty())),
                    setTableStep(
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
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("LL")),
                        new WorkbookMutationAction.EnsureSheet()),
                    new ResolvedStep(
                        new CellSelector.ByAddress(new SheetName("LL"), new CellAddress("E2")),
                        new ResolvedMutationAction.SetComment(
                            new ResolvedComment("Note BudgetTotal", "GridGrind", true))),
                    new ResolvedStep(
                        new CellSelector.ByAddress(new SheetName("LL"), new CellAddress("A2")),
                        new ResolvedMutationAction.SetComment(
                            new ResolvedComment("Note Report_Value", "GridGrind", true))),
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("LL")),
                        new WorkbookMutationAction.EnsureSheet()),
                    new ResolvedStep(
                        new ColumnBandSelector.Span(new SheetName("LL"), 1, 3),
                        new WorkbookMutationAction.DeleteColumns()),
                    new ResolvedStep(
                        new ColumnBandSelector.Span(new SheetName("LL"), 0, 0),
                        new WorkbookMutationAction.DeleteColumns()),
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("LL")),
                        new WorkbookMutationAction.AutoSizeColumns()),
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("LL")),
                        new WorkbookMutationAction.AutoSizeColumns()))));
  }

  @Test
  void requireRoundTripReadableAcceptsCopiedPicturesWithRetargetedRelations() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-copy-picture-roundtrip-",
                List.of(
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("Queue")),
                        new WorkbookMutationAction.EnsureSheet()),
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("Queue")),
                        new ResolvedMutationAction.SetPicture(
                            new ResolvedPicture(
                                "QueuePreview",
                                new BinaryContent(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
                                ExcelPictureFormat.PNG,
                                new DrawingAnchorSpec.TwoCell(
                                    new DrawingMarkerSpec(1, 1, 0, 0),
                                    new DrawingMarkerSpec(4, 6, 0, 0),
                                    dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior
                                        .MOVE_AND_RESIZE),
                                Optional.of("Queue preview")))),
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("Queue")),
                        new ResolvedMutationAction.SetPicture(
                            new ResolvedPicture(
                                "QueuePreview2",
                                new BinaryContent(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
                                ExcelPictureFormat.PNG,
                                new DrawingAnchorSpec.TwoCell(
                                    new DrawingMarkerSpec(6, 1, 0, 0),
                                    new DrawingMarkerSpec(9, 6, 0, 0),
                                    dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior
                                        .MOVE_AND_RESIZE),
                                Optional.empty()))),
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("Queue")),
                        new WorkbookMutationAction.CopySheet(
                            new SheetName("Queue Copy"),
                            new dev.erst.gridgrind.core.vocabulary.SheetCopyPosition
                                .AppendAtEnd())))));
  }

  @Test
  void requireRoundTripReadableAcceptsCopiedEmbeddedObjectPreviewRelations() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-copy-embedded-preview-roundtrip-",
                List.of(
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("Queue")),
                        new WorkbookMutationAction.EnsureSheet()),
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("Queue")),
                        new ResolvedMutationAction.SetEmbeddedObject(
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
                                        .MOVE_AND_RESIZE)))),
                    new ResolvedStep(
                        new CellSelector.ByAddress(new SheetName("Queue"), new CellAddress("B2")),
                        new ResolvedMutationAction.SetComment(
                            new ResolvedComment("Queue note", "GridGrind", true))),
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("Queue")),
                        new WorkbookMutationAction.CopySheet(
                            new SheetName("Queue Copy"),
                            new dev.erst.gridgrind.core.vocabulary.SheetCopyPosition
                                .AppendAtEnd())))));
  }

  @Test
  void requireRoundTripReadableAcceptsArrayFormulasAndSignatureLines() {
    assertDoesNotThrow(
        () ->
            roundTrip(
                "gridgrind-jazzer-array-signature-roundtrip-",
                List.of(
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("Approvals")),
                        new WorkbookMutationAction.EnsureSheet()),
                    new ResolvedStep(
                        new RangeSelector.ByRange(
                            new SheetName("Approvals"), new RangeRef("A1:C4")),
                        new ResolvedMutationAction.SetRange(
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
                                    new CellValue.NumberValue(21.0d))))),
                    new ResolvedStep(
                        new RangeSelector.ByRange(
                            new SheetName("Approvals"), new RangeRef("D2:D4")),
                        new ResolvedMutationAction.SetArrayFormula(
                            new ResolvedArrayFormula("B2:B4*C2:C4"))),
                    new ResolvedStep(
                        new CellSelector.ByAddress(
                            new SheetName("Approvals"), new CellAddress("D2")),
                        new CellMutationAction.ClearArrayFormula()),
                    new ResolvedStep(
                        new SheetSelector.ByName(new SheetName("Approvals")),
                        new ResolvedMutationAction.SetSignatureLine(signatureLineDefinition())))));
  }

  private static void roundTrip(String prefix, List<ResolvedStep> commands) throws IOException {
    Path workbookPath = Files.createTempFile(prefix, ".xlsx");
    Files.deleteIfExists(workbookPath);

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      new WorkbookExecutionEngine().apply(workbook, commands);
      FuzzWorkbookIo.saveWorkbook(workbook, workbookPath);
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
        Optional.of("Review the budget before signing."),
        Optional.of("Ada Lovelace"),
        Optional.of("Finance"),
        Optional.of("ada@example.com"),
        Optional.empty(),
        Optional.of("invalid"),
        Optional.of(ExcelPictureFormat.PNG),
        Optional.of(new BinaryContent(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64))));
  }
}
