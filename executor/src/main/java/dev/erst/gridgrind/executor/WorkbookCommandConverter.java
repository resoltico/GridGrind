package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.ArrayFormulaInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.ExcelArrayFormulaDefinition;
import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelCellAlignment;
import dev.erst.gridgrind.excel.ExcelCellFill;
import dev.erst.gridgrind.excel.ExcelCellFont;
import dev.erst.gridgrind.excel.ExcelCellProtection;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelCustomXmlImportDefinition;
import dev.erst.gridgrind.excel.ExcelCustomXmlMappingLocator;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.ExcelPivotTableDefinition;
import dev.erst.gridgrind.excel.ExcelRichText;
import dev.erst.gridgrind.excel.ExcelRichTextRun;
import dev.erst.gridgrind.excel.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSignatureLineDefinition;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.WorkbookCommand;
import java.util.Optional;

/**
 * Converts contract mutation steps and style inputs into workbook-core commands.
 *
 * <p>This translation seam intentionally spans the full mutation surface on both sides.
 */
@SuppressWarnings("PMD.ExcessiveImports")
final class WorkbookCommandConverter {
  private WorkbookCommandConverter() {}

  /** Converts one protocol mutation step into the matching workbook-core command. */
  static WorkbookCommand toCommand(MutationStep step) {
    return toCommand(step.target(), step.action());
  }

  /**
   * Converts one protocol mutation action plus selector into the matching workbook-core command.
   */
  static WorkbookCommand toCommand(Selector target, MutationAction action) {
    return switch (action) {
      case MutationAction.WorkbookMutationAction workbookAction ->
          WorkbookCommandWorkbookMutationConverter.toCommand(target, workbookAction);
      case MutationAction.CellMutationAction cellAction ->
          WorkbookCommandCellMutationConverter.toCommand(target, cellAction);
      case MutationAction.DrawingMutationAction drawingAction ->
          WorkbookCommandDrawingMutationConverter.toCommand(target, drawingAction);
      case MutationAction.StructuredMutationAction structuredAction ->
          WorkbookCommandStructuredMutationConverter.toCommand(target, structuredAction);
    };
  }

  static ExcelCellValue toExcelCellValue(CellInput value) {
    return WorkbookCommandCellInputConverter.toExcelCellValue(value);
  }

  static ExcelArrayFormulaDefinition toExcelArrayFormulaDefinition(ArrayFormulaInput input) {
    return WorkbookCommandCellInputConverter.toExcelArrayFormulaDefinition(input);
  }

  static ExcelCustomXmlImportDefinition toExcelCustomXmlImportDefinition(
      CustomXmlImportInput input) {
    return WorkbookCommandDrawingInputConverter.toExcelCustomXmlImportDefinition(input);
  }

  static ExcelCustomXmlMappingLocator toExcelCustomXmlMappingLocator(
      dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator locator) {
    return WorkbookCommandDrawingInputConverter.toExcelCustomXmlMappingLocator(locator);
  }

  static ExcelRichText toExcelRichText(CellInput.RichText richText) {
    return WorkbookCommandCellInputConverter.toExcelRichText(richText);
  }

  static ExcelRichTextRun toExcelRichTextRun(RichTextRunInput run) {
    return WorkbookCommandCellInputConverter.toExcelRichTextRun(run);
  }

  static ExcelHyperlink toExcelHyperlink(HyperlinkTarget target) {
    return WorkbookCommandCellInputConverter.toExcelHyperlink(target);
  }

  static ExcelComment toExcelComment(CommentInput comment) {
    return WorkbookCommandCellInputConverter.toExcelComment(comment);
  }

  static ExcelPictureDefinition toExcelPictureDefinition(PictureInput picture) {
    return WorkbookCommandDrawingInputConverter.toExcelPictureDefinition(picture);
  }

  static ExcelChartDefinition toExcelChartDefinition(ChartInput chart) {
    return WorkbookCommandDrawingInputConverter.toExcelChartDefinition(chart);
  }

  static ExcelSignatureLineDefinition toExcelSignatureLineDefinition(
      SignatureLineInput signatureLine) {
    return WorkbookCommandDrawingInputConverter.toExcelSignatureLineDefinition(signatureLine);
  }

  static ExcelShapeDefinition toExcelShapeDefinition(ShapeInput shape) {
    return WorkbookCommandDrawingInputConverter.toExcelShapeDefinition(shape);
  }

  static ExcelEmbeddedObjectDefinition toExcelEmbeddedObjectDefinition(
      EmbeddedObjectInput embeddedObject) {
    return WorkbookCommandDrawingInputConverter.toExcelEmbeddedObjectDefinition(embeddedObject);
  }

  static ExcelDrawingAnchor.TwoCell toExcelDrawingAnchor(DrawingAnchorInput anchor) {
    return WorkbookCommandDrawingInputConverter.toExcelDrawingAnchor(anchor);
  }

  static ExcelCellStyle toExcelCellStyle(CellStyleInput style) {
    return WorkbookCommandCellInputConverter.toExcelCellStyle(style);
  }

  static Optional<ExcelCellAlignment> toExcelCellAlignment(CellAlignmentInput alignment) {
    return WorkbookCommandCellInputConverter.toExcelCellAlignment(alignment);
  }

  static Optional<ExcelCellFont> toExcelCellFont(CellFontInput font) {
    return WorkbookCommandCellInputConverter.toExcelCellFont(font);
  }

  static Optional<ExcelCellFill> toExcelCellFill(CellFillInput fill) {
    return WorkbookCommandCellInputConverter.toExcelCellFill(fill);
  }

  static Optional<ExcelCellProtection> toExcelCellProtection(CellProtectionInput protection) {
    return WorkbookCommandCellInputConverter.toExcelCellProtection(protection);
  }

  static Optional<ExcelFontHeight> toExcelFontHeight(FontHeightInput fontHeight) {
    return WorkbookCommandCellInputConverter.toExcelFontHeight(fontHeight);
  }

  static Optional<ExcelBorder> toExcelBorder(CellBorderInput border) {
    return WorkbookCommandCellInputConverter.toExcelBorder(border);
  }

  static Optional<ExcelBorderSide> toExcelBorderSide(CellBorderSideInput side) {
    return WorkbookCommandCellInputConverter.toExcelBorderSide(side);
  }

  static ExcelDataValidationDefinition toExcelDataValidationDefinition(
      DataValidationInput validation) {
    return WorkbookCommandStructuredInputConverter.toExcelDataValidationDefinition(validation);
  }

  static ExcelDataValidationRule toExcelDataValidationRule(DataValidationRuleInput rule) {
    return WorkbookCommandStructuredInputConverter.toExcelDataValidationRule(rule);
  }

  static Optional<ExcelDataValidationPrompt> toExcelDataValidationPrompt(
      DataValidationPromptInput prompt) {
    return WorkbookCommandStructuredInputConverter.toExcelDataValidationPrompt(prompt);
  }

  static Optional<ExcelDataValidationErrorAlert> toExcelDataValidationErrorAlert(
      DataValidationErrorAlertInput errorAlert) {
    return WorkbookCommandStructuredInputConverter.toExcelDataValidationErrorAlert(errorAlert);
  }

  static ExcelConditionalFormattingBlockDefinition toExcelConditionalFormattingBlock(
      ConditionalFormattingBlockInput block) {
    return WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingBlock(block);
  }

  static ExcelConditionalFormattingRule toExcelConditionalFormattingRule(
      ConditionalFormattingRuleInput rule) {
    return WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(rule);
  }

  static Optional<ExcelDifferentialStyle> toExcelDifferentialStyle(DifferentialStyleInput style) {
    return WorkbookCommandStructuredInputConverter.toExcelDifferentialStyle(style);
  }

  static Optional<ExcelDifferentialBorder> toExcelDifferentialBorder(
      DifferentialBorderInput border) {
    return WorkbookCommandStructuredInputConverter.toExcelDifferentialBorder(border);
  }

  static ExcelTableDefinition toExcelTableDefinition(TableInput table) {
    return WorkbookCommandTabularInputConverter.toExcelTableDefinition(table);
  }

  static ExcelPivotTableDefinition toExcelPivotTableDefinition(PivotTableInput pivotTable) {
    return WorkbookCommandTabularInputConverter.toExcelPivotTableDefinition(pivotTable);
  }

  static ExcelTableStyle toExcelTableStyle(TableStyleInput style) {
    return WorkbookCommandTabularInputConverter.toExcelTableStyle(style);
  }

  static ExcelNamedRangeScope toExcelNamedRangeScope(NamedRangeScope scope) {
    return WorkbookCommandLayoutInputConverter.toExcelNamedRangeScope(scope);
  }

  static ExcelNamedRangeScope toExcelNamedRangeScope(NamedRangeSelector.ScopedExact selector) {
    return WorkbookCommandLayoutInputConverter.toExcelNamedRangeScope(selector);
  }

  static String toExcelNamedRangeName(NamedRangeSelector.ScopedExact selector) {
    return WorkbookCommandLayoutInputConverter.toExcelNamedRangeName(selector);
  }

  static ExcelNamedRangeTarget toExcelNamedRangeTarget(NamedRangeTarget target) {
    return WorkbookCommandLayoutInputConverter.toExcelNamedRangeTarget(target);
  }

  /* Converts a protocol sheet copy-position variant into the engine position type. */
  static ExcelSheetCopyPosition toExcelSheetCopyPosition(SheetCopyPosition position) {
    return WorkbookCommandLayoutInputConverter.toExcelSheetCopyPosition(position);
  }
}
