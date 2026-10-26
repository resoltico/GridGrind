package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceSelectorSupport.*;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;

import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.source.*;
import dev.erst.gridgrind.core.vocabulary.*;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelComparisonFormulaSupport;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.Optional;
import java.util.stream.Stream;

/** Generates rule-oriented authored and engine values for operation-sequence fuzzing. */
final class OperationSequenceRuleValues {
  private OperationSequenceRuleValues() {}

  static DataValidationSpec nextDataValidationInput(GridGrindFuzzData data) {
    return new DataValidationSpec(
        data.consumeBoolean()
            ? new DataValidationRuleSpec.ExplicitList(nextExplicitValidationValues(data))
            : new DataValidationRuleSpec.WholeNumber(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "1", Optional.empty()),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean()
            ? Optional.of(
                new DataValidationPromptSpec(
                    TextSourceSpec.inline("Status"),
                    TextSourceSpec.inline("Use an allowed value."),
                    data.consumeBoolean()))
            : Optional.empty(),
        data.consumeBoolean()
            ? Optional.of(
                new DataValidationErrorAlertSpec(
                    ExcelDataValidationErrorStyle.STOP,
                    TextSourceSpec.inline("Invalid"),
                    TextSourceSpec.inline("Use an allowed value."),
                    data.consumeBoolean()))
            : Optional.empty());
  }

  static ResolvedDataValidation nextExcelDataValidationDefinition(GridGrindFuzzData data) {
    return new ResolvedDataValidation(
        data.consumeBoolean()
            ? new DataValidationRuleSpec.ExplicitList(List.of("Queued", "Done"))
            : new DataValidationRuleSpec.WholeNumber(
                ExcelComparisonOperator.GREATER_OR_EQUAL,
                ExcelComparisonFormulaSupport.normalizeFormula("1", "formula1"),
                ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                    Optional.empty(), "formula2")),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean()
            ? Optional.of(
                new ResolvedValidationPrompt(
                    "Status", "Use an allowed value.", data.consumeBoolean()))
            : Optional.empty(),
        data.consumeBoolean()
            ? Optional.of(
                new ResolvedValidationErrorAlert(
                    ExcelDataValidationErrorStyle.STOP,
                    "Invalid",
                    "Use an allowed value.",
                    data.consumeBoolean()))
            : Optional.empty());
  }

  private static List<String> nextExplicitValidationValues(GridGrindFuzzData data) {
    return data.consumeBoolean() ? List.of("Queued", "Done") : List.of("x".repeat(253));
  }

  static ConditionalFormattingSpec nextConditionalFormattingInput(GridGrindFuzzData data) {
    return new ConditionalFormattingSpec(
        List.of(
            data.consumeBoolean()
                ? new ConditionalFormattingRuleSpec.FormulaRule(
                    "A1>0", data.consumeBoolean(), Optional.of(nextDifferentialStyleInput(data)))
                : new ConditionalFormattingRuleSpec.CellValueRule(
                    ExcelComparisonOperator.GREATER_THAN,
                    "1",
                    Optional.empty(),
                    data.consumeBoolean(),
                    Optional.of(nextDifferentialStyleInput(data)))));
  }

  static ResolvedStep nextConditionalFormattingCommand(
      GridGrindFuzzData data, boolean validRange, String sheetName) {
    return new ResolvedStep(
        new RangeSelector.ByRanges(
            new SheetName(sheetName),
            List.of(
                new dev.erst.gridgrind.core.vocabulary.RangeRef(
                    validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)))),
        new StructuredMutationAction.SetConditionalFormatting(
            new ConditionalFormattingSpec(
                List.of(
                    data.consumeBoolean()
                        ? new ConditionalFormattingRuleSpec.FormulaRule(
                            "A1>0",
                            data.consumeBoolean(),
                            Optional.of(nextExcelDifferentialStyle(data)))
                        : new ConditionalFormattingRuleSpec.CellValueRule(
                            ExcelComparisonOperator.GREATER_THAN,
                            "1",
                            Optional.empty(),
                            data.consumeBoolean(),
                            Optional.of(nextExcelDifferentialStyle(data)))))));
  }

  static DifferentialStylePatch nextDifferentialStyleInput(GridGrindFuzzData data) {
    boolean includeNumberFormat = data.consumeBoolean();
    Boolean bold = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean italic = data.consumeBoolean() ? Boolean.TRUE : null;
    ColorSpec fontColor = data.consumeBoolean() ? nextColorInput(data) : null;
    Boolean underline = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean strikeout = data.consumeBoolean() ? Boolean.TRUE : null;
    ColorSpec fillColor = data.consumeBoolean() ? nextColorInput(data) : null;
    String numberFormat =
        includeNumberFormat
                || Stream.of(bold, italic, fontColor, underline, strikeout, fillColor)
                    .allMatch(Objects::isNull)
            ? "0.00"
            : null;
    return new DifferentialStylePatch(
        Optional.ofNullable(numberFormat),
        Optional.ofNullable(bold),
        Optional.ofNullable(italic),
        Optional.empty(),
        java.util.Optional.ofNullable(fontColor),
        Optional.ofNullable(underline),
        Optional.ofNullable(strikeout),
        java.util.Optional.ofNullable(fillColor),
        java.util.Optional.empty());
  }

  static DifferentialStylePatch nextExcelDifferentialStyle(GridGrindFuzzData data) {
    boolean includeNumberFormat = data.consumeBoolean();
    Boolean bold = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean italic = data.consumeBoolean() ? Boolean.TRUE : null;
    ColorSpec fontColor = data.consumeBoolean() ? nextDifferentialColor(data) : null;
    Boolean underline = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean strikeout = data.consumeBoolean() ? Boolean.TRUE : null;
    ColorSpec fillColor = data.consumeBoolean() ? nextDifferentialColor(data) : null;
    String numberFormat =
        includeNumberFormat
                || Stream.of(bold, italic, fontColor, underline, strikeout, fillColor)
                    .allMatch(Objects::isNull)
            ? "0.00"
            : null;
    return new DifferentialStylePatch(
        Optional.ofNullable(numberFormat),
        Optional.ofNullable(bold),
        Optional.ofNullable(italic),
        Optional.empty(),
        Optional.ofNullable(fontColor),
        Optional.ofNullable(underline),
        Optional.ofNullable(strikeout),
        Optional.ofNullable(fillColor),
        Optional.empty());
  }

  private static ColorSpec nextColorInput(GridGrindFuzzData data) {
    Optional<Double> tint = nextTint(data);
    return switch (Math.floorMod(data.consumeByte(), 3)) {
      case 0 -> new ColorSpec.Rgb(new RgbHex("#102030"), tint);
      case 1 -> new ColorSpec.Theme(data.consumeInt(0, 11), tint);
      default -> new ColorSpec.Indexed(data.consumeInt(0, 63), tint);
    };
  }

  private static ColorSpec nextDifferentialColor(GridGrindFuzzData data) {
    Optional<Double> tint = nextTint(data);
    return switch (Math.floorMod(data.consumeByte(), 3)) {
      case 0 -> new ColorSpec.Rgb(new RgbHex("#E0F0AA"), tint);
      case 1 -> new ColorSpec.Theme(data.consumeInt(0, 11), tint);
      default -> new ColorSpec.Indexed(data.consumeInt(0, 63), tint);
    };
  }

  private static Optional<Double> nextTint(GridGrindFuzzData data) {
    return data.consumeBoolean()
        ? Optional.of(data.consumeRegularDouble(-1.0d, 1.0d))
        : Optional.empty();
  }

  static RangeSelector nextRangeSelector(
      GridGrindFuzzData data, String sheetName, boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new RangeSelector.AllOnSheet(new SheetName(sheetName));
      default ->
          new RangeSelector.ByRanges(
              new SheetName(sheetName),
              List.of(
                  new RangeRef(
                      validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false))));
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
