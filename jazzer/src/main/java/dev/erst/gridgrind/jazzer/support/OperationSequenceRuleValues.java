package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceSelectorSupport.*;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.source.*;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.foundation.*;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.stream.Stream;

/** Generates rule-oriented authored and engine values for operation-sequence fuzzing. */
final class OperationSequenceRuleValues {
  private OperationSequenceRuleValues() {}

  static DataValidationInput nextDataValidationInput(GridGrindFuzzData data) {
    return new DataValidationInput(
        data.consumeBoolean()
            ? new DataValidationRuleInput.ExplicitList(List.of("Queued", "Done"))
            : new DataValidationRuleInput.WholeNumber(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean()
            ? new DataValidationPromptInput(
                TextSourceInput.inline("Status"),
                TextSourceInput.inline("Use an allowed value."),
                data.consumeBoolean())
            : null,
        data.consumeBoolean()
            ? new DataValidationErrorAlertInput(
                ExcelDataValidationErrorStyle.STOP,
                TextSourceInput.inline("Invalid"),
                TextSourceInput.inline("Use an allowed value."),
                data.consumeBoolean())
            : null);
  }

  static ExcelDataValidationDefinition nextExcelDataValidationDefinition(GridGrindFuzzData data) {
    return new ExcelDataValidationDefinition(
        data.consumeBoolean()
            ? new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done"))
            : new ExcelDataValidationRule.WholeNumber(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean()
            ? new ExcelDataValidationPrompt(
                "Status", "Use an allowed value.", data.consumeBoolean())
            : null,
        data.consumeBoolean()
            ? new ExcelDataValidationErrorAlert(
                ExcelDataValidationErrorStyle.STOP,
                "Invalid",
                "Use an allowed value.",
                data.consumeBoolean())
            : null);
  }

  static ConditionalFormattingBlockInput nextConditionalFormattingInput(
      GridGrindFuzzData data, boolean validRange) {
    return new ConditionalFormattingBlockInput(
        List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)),
        List.of(
            data.consumeBoolean()
                ? new ConditionalFormattingRuleInput.FormulaRule(
                    "A1>0", data.consumeBoolean(), nextDifferentialStyleInput(data))
                : new ConditionalFormattingRuleInput.CellValueRule(
                    ExcelComparisonOperator.GREATER_THAN,
                    "1",
                    null,
                    data.consumeBoolean(),
                    nextDifferentialStyleInput(data))));
  }

  static ExcelConditionalFormattingBlockDefinition nextExcelConditionalFormattingBlockDefinition(
      GridGrindFuzzData data, boolean validRange) {
    return new ExcelConditionalFormattingBlockDefinition(
        List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)),
        List.of(
            data.consumeBoolean()
                ? new ExcelConditionalFormattingRule.FormulaRule(
                    "A1>0", data.consumeBoolean(), nextExcelDifferentialStyle(data))
                : new ExcelConditionalFormattingRule.CellValueRule(
                    ExcelComparisonOperator.GREATER_THAN,
                    "1",
                    null,
                    data.consumeBoolean(),
                    nextExcelDifferentialStyle(data))));
  }

  static DifferentialStyleInput nextDifferentialStyleInput(GridGrindFuzzData data) {
    boolean includeNumberFormat = data.consumeBoolean();
    Boolean bold = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean italic = data.consumeBoolean() ? Boolean.TRUE : null;
    String fontColor = data.consumeBoolean() ? "#102030" : null;
    Boolean underline = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean strikeout = data.consumeBoolean() ? Boolean.TRUE : null;
    String fillColor = data.consumeBoolean() ? "#E0F0AA" : null;
    String numberFormat =
        includeNumberFormat
                || Stream.of(bold, italic, fontColor, underline, strikeout, fillColor)
                    .allMatch(Objects::isNull)
            ? "0.00"
            : null;
    return new DifferentialStyleInput(
        numberFormat,
        bold,
        italic,
        null,
        java.util.Optional.ofNullable(fontColor),
        underline,
        strikeout,
        java.util.Optional.ofNullable(fillColor),
        java.util.Optional.empty());
  }

  static ExcelDifferentialStyle nextExcelDifferentialStyle(GridGrindFuzzData data) {
    boolean includeNumberFormat = data.consumeBoolean();
    Boolean bold = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean italic = data.consumeBoolean() ? Boolean.TRUE : null;
    String fontColor = data.consumeBoolean() ? "#102030" : null;
    Boolean underline = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean strikeout = data.consumeBoolean() ? Boolean.TRUE : null;
    String fillColor = data.consumeBoolean() ? "#E0F0AA" : null;
    String numberFormat =
        includeNumberFormat
                || Stream.of(bold, italic, fontColor, underline, strikeout, fillColor)
                    .allMatch(Objects::isNull)
            ? "0.00"
            : null;
    return new ExcelDifferentialStyle(
        numberFormat, bold, italic, null, fontColor, underline, strikeout, fillColor, null);
  }

  static RangeSelector nextRangeSelector(
      GridGrindFuzzData data, String sheetName, boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new RangeSelector.AllOnSheet(sheetName);
      default ->
          new RangeSelector.ByRanges(
              sheetName,
              List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)));
    };
  }

  static ExcelRangeSelection nextExcelRangeSelection(GridGrindFuzzData data, boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new ExcelRangeSelection.All();
      default ->
          new ExcelRangeSelection.Selected(
              List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)));
    };
  }

  static String nextAutofilterRange(boolean validRange) {
    return validRange ? "E1:F3" : "BadRange";
  }

  static String nextCopySheetName(String sourceSheetName) {
    Objects.requireNonNull(sourceSheetName, "sourceSheetName must not be null");
    String base =
        sourceSheetName.length() <= 27 ? sourceSheetName : sourceSheetName.substring(0, 27);
    return base + "_B1";
  }
}
