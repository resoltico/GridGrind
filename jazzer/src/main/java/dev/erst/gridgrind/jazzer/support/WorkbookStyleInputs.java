package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import java.math.BigDecimal;
import java.util.Objects;
import java.util.Optional;

/** Creates bounded protocol style patches for fuzz harnesses. */
public final class WorkbookStyleInputs {
  private WorkbookStyleInputs() {}

  /** Returns a bounded protocol style patch. */
  public static CellStylePatch nextStyleInput(GridGrindFuzzData data) {
    Objects.requireNonNull(data, "data must not be null");

    return switch (data.consumeInt(0, 7)) {
      case 0 ->
          new CellStylePatch(
              Optional.of(data.consumeBoolean() ? "0.00" : "yyyy-mm-dd"),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 1 ->
          new CellStylePatch(
              Optional.empty(),
              Optional.of(nextAlignmentInput(data, true)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 2 ->
          new CellStylePatch(
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFontInput(data, true)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 3 ->
          new CellStylePatch(
              Optional.empty(),
              Optional.of(nextAlignmentInput(data, true)),
              Optional.of(nextFontInput(data, true)),
              Optional.of(nextFillInput(data)),
              Optional.of(nextProtocolBorder(data)),
              Optional.ofNullable(nextProtectionInput(data)));
      case 4 ->
          new CellStylePatch(
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFontInput(data, false)),
              Optional.of(nextFillInput(data)),
              Optional.of(nextProtocolBorder(data)),
              Optional.empty());
      case 5 ->
          new CellStylePatch(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFillInput(data)),
              Optional.empty(),
              Optional.ofNullable(nextProtectionInput(data)));
      case 6 ->
          new CellStylePatch(
              Optional.empty(),
              Optional.of(nextAlignmentInput(data, true)),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextProtocolBorder(data)),
              Optional.empty());
      default ->
          new CellStylePatch(
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFontInput(data, false)),
              Optional.empty(),
              Optional.empty(),
              Optional.ofNullable(nextProtectionInput(data)));
    };
  }

  private static CellAlignmentPatch nextAlignmentInput(
      GridGrindFuzzData data, boolean includeDepth) {
    return new CellAlignmentPatch(
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

  private static CellFontPatch nextFontInput(GridGrindFuzzData data, boolean includeName) {
    return new CellFontPatch(
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

  private static CellFillSpec nextFillInput(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 3)) {
      case 0 -> CellFillSpec.patternForeground(ExcelFillPattern.SOLID, nextColorInput(data));
      case 1 ->
          data.consumeBoolean()
              ? CellFillSpec.patternColors(
                  nextPatternFill(data), nextColorInput(data), nextColorInput(data))
              : CellFillSpec.patternForeground(nextPatternFill(data), nextColorInput(data));
      case 2 -> CellFillSpec.patternBackground(nextPatternFill(data), nextColorInput(data));
      default ->
          CellFillSpec.pattern(
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

  private static FontHeightSpec nextFontHeightInput(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 2)) {
      case 0 -> new FontHeightSpec.Twips(data.consumeInt(20, 640));
      case 1 -> new FontHeightSpec.Points(nextPointHeight(data));
      default -> new FontHeightSpec.Twips(220);
    };
  }

  private static BigDecimal nextPointHeight(GridGrindFuzzData data) {
    return new ExcelFontHeight(data.consumeInt(20, 640)).points();
  }

  private static String nextRgbHex(GridGrindFuzzData data) {
    return "#%02X%02X%02X"
        .formatted(data.consumeInt(0, 255), data.consumeInt(0, 255), data.consumeInt(0, 255));
  }

  private static ColorSpec nextColorInput(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 2)) {
      case 0 ->
          data.consumeBoolean()
              ? ColorSpec.rgb(new RgbHex(nextRgbHex(data)), nextTint(data))
              : ColorSpec.rgb(new RgbHex(nextRgbHex(data)));
      case 1 ->
          data.consumeBoolean()
              ? ColorSpec.theme(data.consumeInt(0, 9), nextTint(data))
              : ColorSpec.theme(data.consumeInt(0, 9));
      default ->
          data.consumeBoolean()
              ? ColorSpec.indexed(data.consumeInt(0, 64), nextTint(data))
              : ColorSpec.indexed(data.consumeInt(0, 64));
    };
  }

  private static CellBorderPatch nextProtocolBorder(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 4)) {
      case 0 ->
          new CellBorderPatch(
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 1 ->
          new CellBorderPatch(
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.empty());
      case 2 ->
          new CellBorderPatch(
              Optional.empty(),
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextBorderSide(data)));
      case 3 ->
          new CellBorderPatch(
              Optional.of(nextBorderSide(data)),
              Optional.of(nextBorderSide(data)),
              Optional.of(nextBorderSide(data)),
              Optional.of(nextBorderSide(data)),
              Optional.of(nextBorderSide(data)));
      default ->
          new CellBorderPatch(
              Optional.empty(),
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.of(nextBorderSide(data)),
              Optional.empty());
    };
  }

  private static BorderSidePatch nextBorderSide(GridGrindFuzzData data) {
    ExcelBorderStyle[] values = ExcelBorderStyle.values();
    ExcelBorderStyle style = values[data.consumeInt(0, values.length - 1)];
    return new BorderSidePatch(
        Optional.of(style),
        style == ExcelBorderStyle.NONE || !data.consumeBoolean()
            ? Optional.empty()
            : Optional.of(nextColorInput(data)));
  }

  private static CellProtectionPatch nextProtectionInput(GridGrindFuzzData data) {
    return new CellProtectionPatch(
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
