package dev.erst.gridgrind.fuzz.support;

import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellGradientFillSpec;
import dev.erst.gridgrind.core.model.style.CellGradientStopSpec;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Owns bounded style, color, border, and font decoding for fuzz inputs. */
final class FuzzStyleDecoders {
  private FuzzStyleDecoders() {}

  static CellStylePatch nextStyle(GridGrindFuzzData data) {
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
              Optional.of(nextAlignment(data, true)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 2 ->
          new CellStylePatch(
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFont(data, true)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 3 ->
          new CellStylePatch(
              Optional.empty(),
              Optional.of(nextAlignment(data, true)),
              Optional.of(nextFont(data, true)),
              Optional.of(nextFill(data)),
              Optional.of(nextBorderPatch(data)),
              Optional.ofNullable(nextProtection(data)));
      case 4 ->
          new CellStylePatch(
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFont(data, false)),
              Optional.of(nextFill(data)),
              Optional.of(nextBorderPatch(data)),
              Optional.empty());
      case 5 ->
          new CellStylePatch(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFill(data)),
              Optional.empty(),
              Optional.ofNullable(nextProtection(data)));
      case 6 ->
          new CellStylePatch(
              Optional.empty(),
              Optional.of(nextAlignment(data, true)),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextBorderPatch(data)),
              Optional.empty());
      default ->
          new CellStylePatch(
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFont(data, false)),
              Optional.empty(),
              Optional.empty(),
              Optional.ofNullable(nextProtection(data)));
    };
  }

  static CellFontPatch nextRichTextFontPatch(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 6)) {
      case 0 ->
          new CellFontPatch(
              Optional.of(Boolean.TRUE),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 1 ->
          new CellFontPatch(
              Optional.empty(),
              Optional.of(Boolean.FALSE),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 2 ->
          new CellFontPatch(
              Optional.empty(),
              Optional.empty(),
              Optional.of("Aptos"),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 3 ->
          new CellFontPatch(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFontHeightSpec(data)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 4 ->
          new CellFontPatch(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextColor(data)),
              Optional.empty(),
              Optional.empty());
      case 5 ->
          new CellFontPatch(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(Boolean.TRUE),
              Optional.empty());
      default ->
          new CellFontPatch(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(Boolean.FALSE));
    };
  }

  private static ExcelHorizontalAlignment nextHorizontalAlignment(GridGrindFuzzData data) {
    ExcelHorizontalAlignment[] values = ExcelHorizontalAlignment.values();
    return values[data.consumeInt(0, values.length - 1)];
  }

  private static CellAlignmentPatch nextAlignment(GridGrindFuzzData data, boolean includeDepth) {
    return new CellAlignmentPatch(
        data.consumeBoolean() ? Optional.of(Boolean.TRUE) : Optional.empty(),
        Optional.of(nextHorizontalAlignment(data)),
        Optional.of(nextVerticalAlignment(data)),
        includeDepth && data.consumeBoolean()
            ? Optional.of(data.consumeInt(0, 180))
            : Optional.empty(),
        includeDepth && data.consumeBoolean()
            ? Optional.of(data.consumeInt(0, 8))
            : Optional.empty());
  }

  private static CellFontPatch nextFont(GridGrindFuzzData data, boolean includeName) {
    return new CellFontPatch(
        data.consumeBoolean() ? Optional.of(Boolean.TRUE) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(Boolean.FALSE) : Optional.empty(),
        includeName
            ? Optional.of(data.consumeBoolean() ? "Aptos" : "Aptos Display")
            : Optional.empty(),
        data.consumeBoolean() ? Optional.of(nextFontHeightSpec(data)) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(nextColor(data)) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(Boolean.TRUE) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(Boolean.FALSE) : Optional.empty());
  }

  private static CellFillSpec nextFill(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 4)) {
      case 0 -> CellFillSpec.patternForeground(ExcelFillPattern.SOLID, nextColor(data));
      case 1 ->
          data.consumeBoolean()
              ? CellFillSpec.patternColors(nextPatternFill(data), nextColor(data), nextColor(data))
              : CellFillSpec.patternForeground(nextPatternFill(data), nextColor(data));
      case 2 -> CellFillSpec.patternBackground(nextPatternFill(data), nextColor(data));
      case 3 -> CellFillSpec.gradient(nextGradientFill(data));
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

  private static CellProtectionPatch nextProtection(GridGrindFuzzData data) {
    return new CellProtectionPatch(
        data.consumeBoolean() ? Optional.of(data.consumeBoolean()) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(data.consumeBoolean()) : Optional.empty());
  }

  private static ExcelVerticalAlignment nextVerticalAlignment(GridGrindFuzzData data) {
    ExcelVerticalAlignment[] values = ExcelVerticalAlignment.values();
    return values[data.consumeInt(0, values.length - 1)];
  }

  private static FontHeightSpec nextFontHeightSpec(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 2)) {
      case 0, 1 -> new FontHeightSpec.Twips(data.consumeInt(20, 640));
      default -> new FontHeightSpec.Twips(220);
    };
  }

  private static String nextRgbHex(GridGrindFuzzData data) {
    return "#%02X%02X%02X"
        .formatted(data.consumeInt(0, 255), data.consumeInt(0, 255), data.consumeInt(0, 255));
  }

  private static ColorSpec nextColor(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 2)) {
      case 0 ->
          new ColorSpec.Rgb(
              new RgbHex(nextRgbHex(data)),
              data.consumeBoolean() ? Optional.of(nextTint(data)) : Optional.empty());
      case 1 ->
          new ColorSpec.Theme(
              data.consumeInt(0, 9),
              data.consumeBoolean() ? Optional.of(nextTint(data)) : Optional.empty());
      default ->
          new ColorSpec.Indexed(
              data.consumeInt(0, 64),
              data.consumeBoolean() ? Optional.of(nextTint(data)) : Optional.empty());
    };
  }

  private static CellGradientFillSpec nextGradientFill(GridGrindFuzzData data) {
    boolean linear = data.consumeBoolean();
    List<CellGradientStopSpec> stops =
        List.of(
            new CellGradientStopSpec(0.0d, nextColor(data)),
            new CellGradientStopSpec(1.0d, nextColor(data)));
    if (linear) {
      return CellGradientFillSpec.linear(
          data.consumeBoolean()
              ? Optional.of(data.consumeRegularDouble(0.0d, 180.0d))
              : Optional.empty(),
          stops);
    }
    return CellGradientFillSpec.path(
        data.consumeBoolean()
            ? Optional.of(data.consumeRegularDouble(0.0d, 1.0d))
            : Optional.empty(),
        data.consumeBoolean()
            ? Optional.of(data.consumeRegularDouble(0.0d, 1.0d))
            : Optional.empty(),
        data.consumeBoolean()
            ? Optional.of(data.consumeRegularDouble(0.0d, 1.0d))
            : Optional.empty(),
        data.consumeBoolean()
            ? Optional.of(data.consumeRegularDouble(0.0d, 1.0d))
            : Optional.empty(),
        stops);
  }

  private static double nextTint(GridGrindFuzzData data) {
    return data.consumeRegularDouble(-1.0d, 1.0d);
  }

  private static CellBorderPatch nextBorderPatch(GridGrindFuzzData data) {
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
            : Optional.of(nextColor(data)));
  }
}
