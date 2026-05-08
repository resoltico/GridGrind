package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.Objects;
import java.util.Optional;

/** Creates bounded protocol style patches for fuzz harnesses. */
public final class WorkbookStyleInputs {
  private WorkbookStyleInputs() {}

  /** Returns a bounded protocol style patch. */
  public static CellStyleInput nextStyleInput(GridGrindFuzzData data) {
    Objects.requireNonNull(data, "data must not be null");

    return switch (data.consumeInt(0, 7)) {
      case 0 ->
          new CellStyleInput(
              Optional.of(data.consumeBoolean() ? "0.00" : "yyyy-mm-dd"),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 1 ->
          new CellStyleInput(
              Optional.empty(),
              Optional.of(nextAlignmentInput(data, true)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 2 ->
          new CellStyleInput(
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFontInput(data, true)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 3 ->
          new CellStyleInput(
              Optional.empty(),
              Optional.of(nextAlignmentInput(data, true)),
              Optional.of(nextFontInput(data, true)),
              Optional.of(nextFillInput(data)),
              Optional.of(nextProtocolBorder(data)),
              Optional.ofNullable(nextProtectionInput(data)));
      case 4 ->
          new CellStyleInput(
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFontInput(data, false)),
              Optional.of(nextFillInput(data)),
              Optional.of(nextProtocolBorder(data)),
              Optional.empty());
      case 5 ->
          new CellStyleInput(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFillInput(data)),
              Optional.empty(),
              Optional.ofNullable(nextProtectionInput(data)));
      case 6 ->
          new CellStyleInput(
              Optional.empty(),
              Optional.of(nextAlignmentInput(data, true)),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextProtocolBorder(data)),
              Optional.empty());
      default ->
          new CellStyleInput(
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFontInput(data, false)),
              Optional.empty(),
              Optional.empty(),
              Optional.ofNullable(nextProtectionInput(data)));
    };
  }

  private static CellAlignmentInput nextAlignmentInput(
      GridGrindFuzzData data, boolean includeDepth) {
    return new CellAlignmentInput(
        data.consumeBoolean() ? Optional.of(Boolean.TRUE) : Optional.empty(),
        Optional.ofNullable(nextHorizontalAlignment(data)),
        Optional.ofNullable(nextVerticalAlignment(data)),
        includeDepth && data.consumeBoolean()
            ? Optional.of(data.consumeInt(0, 180))
            : Optional.empty(),
        includeDepth && data.consumeBoolean()
            ? Optional.of(data.consumeInt(0, 8))
            : Optional.empty());
  }

  private static CellFontInput nextFontInput(GridGrindFuzzData data, boolean includeName) {
    return new CellFontInput(
        data.consumeBoolean() ? Optional.of(Boolean.TRUE) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(Boolean.FALSE) : Optional.empty(),
        includeName
            ? Optional.of(data.consumeBoolean() ? "Aptos" : "Aptos Display")
            : Optional.empty(),
        data.consumeBoolean() ? Optional.of(nextFontHeightInput(data)) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(nextColorInput(data)) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(Boolean.TRUE) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(Boolean.FALSE) : Optional.empty());
  }

  private static CellFillInput nextFillInput(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 3)) {
      case 0 -> CellFillInput.patternForeground(ExcelFillPattern.SOLID, nextColorInput(data));
      case 1 ->
          data.consumeBoolean()
              ? CellFillInput.patternColors(
                  nextPatternFill(data), nextColorInput(data), nextColorInput(data))
              : CellFillInput.patternForeground(nextPatternFill(data), nextColorInput(data));
      case 2 -> CellFillInput.patternBackground(nextPatternFill(data), nextColorInput(data));
      default ->
          CellFillInput.pattern(
              data.consumeBoolean() ? ExcelFillPattern.NONE : nextPatternFill(data));
    };
  }

  private static ExcelFillPattern nextPatternFill(GridGrindFuzzData data) {
    ExcelFillPattern[] patterns =
        new ExcelFillPattern[] {
          ExcelFillPattern.FINE_DOTS,
          ExcelFillPattern.SPARSE_DOTS,
          ExcelFillPattern.THIN_HORIZONTAL_BANDS,
          ExcelFillPattern.THICK_FORWARD_DIAGONAL,
          ExcelFillPattern.SQUARES
        };
    return patterns[data.consumeInt(0, patterns.length - 1)];
  }

  private static FontHeightInput nextFontHeightInput(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 2)) {
      case 0 -> new FontHeightInput.Twips(data.consumeInt(20, 640));
      case 1 -> new FontHeightInput.Points(nextPointHeight(data));
      default -> new FontHeightInput.Twips(220);
    };
  }

  private static BigDecimal nextPointHeight(GridGrindFuzzData data) {
    return new ExcelFontHeight(data.consumeInt(20, 640)).points();
  }

  private static String nextRgbHex(GridGrindFuzzData data) {
    return "#%02X%02X%02X"
        .formatted(data.consumeInt(0, 255), data.consumeInt(0, 255), data.consumeInt(0, 255));
  }

  private static ColorInput nextColorInput(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 2)) {
      case 0 ->
          data.consumeBoolean()
              ? ColorInput.rgb(nextRgbHex(data), nextTint(data))
              : ColorInput.rgb(nextRgbHex(data));
      case 1 ->
          data.consumeBoolean()
              ? ColorInput.theme(data.consumeInt(0, 9), nextTint(data))
              : ColorInput.theme(data.consumeInt(0, 9));
      default ->
          data.consumeBoolean()
              ? ColorInput.indexed(data.consumeInt(0, 64), nextTint(data))
              : ColorInput.indexed(data.consumeInt(0, 64));
    };
  }

  private static CellBorderInput nextProtocolBorder(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 4)) {
      case 0 ->
          new CellBorderInput(
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 1 ->
          new CellBorderInput(
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.empty());
      case 2 ->
          new CellBorderInput(
              Optional.empty(),
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextBorderSide(data)));
      case 3 ->
          new CellBorderInput(
              Optional.of(nextBorderSide(data)),
              Optional.of(nextBorderSide(data)),
              Optional.of(nextBorderSide(data)),
              Optional.of(nextBorderSide(data)),
              Optional.of(nextBorderSide(data)));
      default ->
          new CellBorderInput(
              Optional.empty(),
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.of(nextBorderSide(data)),
              Optional.empty());
    };
  }

  private static CellBorderSideInput nextBorderSide(GridGrindFuzzData data) {
    ExcelBorderStyle[] values = ExcelBorderStyle.values();
    ExcelBorderStyle style = values[data.consumeInt(0, values.length - 1)];
    return new CellBorderSideInput(
        Optional.of(style),
        style == ExcelBorderStyle.NONE || !data.consumeBoolean()
            ? Optional.empty()
            : Optional.of(nextColorInput(data)));
  }

  private static CellProtectionInput nextProtectionInput(GridGrindFuzzData data) {
    return new CellProtectionInput(
        data.consumeBoolean() ? Optional.of(data.consumeBoolean()) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(data.consumeBoolean()) : Optional.empty());
  }

  private static ExcelHorizontalAlignment nextHorizontalAlignment(GridGrindFuzzData data) {
    ExcelHorizontalAlignment[] values = ExcelHorizontalAlignment.values();
    return values[data.consumeInt(0, values.length - 1)];
  }

  private static ExcelVerticalAlignment nextVerticalAlignment(GridGrindFuzzData data) {
    ExcelVerticalAlignment[] values = ExcelVerticalAlignment.values();
    return values[data.consumeInt(0, values.length - 1)];
  }

  private static double nextTint(GridGrindFuzzData data) {
    return data.consumeRegularDouble(-1.0d, 1.0d);
  }
}
