package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelCellAlignment;
import dev.erst.gridgrind.excel.ExcelCellFill;
import dev.erst.gridgrind.excel.ExcelCellFont;
import dev.erst.gridgrind.excel.ExcelCellProtection;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelGradientFill;
import dev.erst.gridgrind.excel.ExcelGradientStop;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Owns bounded style, color, border, and font decoding for Jazzer inputs. */
final class FuzzStyleDecoders {
  private FuzzStyleDecoders() {}

  static ExcelCellStyle nextStyle(GridGrindFuzzData data) {
    Objects.requireNonNull(data, "data must not be null");

    return switch (data.consumeInt(0, 7)) {
      case 0 ->
          new ExcelCellStyle(
              Optional.of(data.consumeBoolean() ? "0.00" : "yyyy-mm-dd"),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 1 ->
          new ExcelCellStyle(
              Optional.empty(),
              Optional.of(nextAlignment(data, true)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 2 ->
          new ExcelCellStyle(
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFont(data, true)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 3 ->
          new ExcelCellStyle(
              Optional.empty(),
              Optional.of(nextAlignment(data, true)),
              Optional.of(nextFont(data, true)),
              Optional.of(nextFill(data)),
              Optional.of(nextExcelBorder(data)),
              Optional.ofNullable(nextProtection(data)));
      case 4 ->
          new ExcelCellStyle(
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFont(data, false)),
              Optional.of(nextFill(data)),
              Optional.of(nextExcelBorder(data)),
              Optional.empty());
      case 5 ->
          new ExcelCellStyle(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFill(data)),
              Optional.empty(),
              Optional.ofNullable(nextProtection(data)));
      case 6 ->
          new ExcelCellStyle(
              Optional.empty(),
              Optional.of(nextAlignment(data, true)),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextExcelBorder(data)),
              Optional.empty());
      default ->
          new ExcelCellStyle(
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextFont(data, false)),
              Optional.empty(),
              Optional.empty(),
              Optional.ofNullable(nextProtection(data)));
    };
  }

  static ExcelCellFont nextRichTextFontPatch(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 6)) {
      case 0 ->
          new ExcelCellFont(
              Optional.of(Boolean.TRUE),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 1 ->
          new ExcelCellFont(
              Optional.empty(),
              Optional.of(Boolean.FALSE),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 2 ->
          new ExcelCellFont(
              Optional.empty(),
              Optional.empty(),
              Optional.of("Aptos"),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 3 ->
          new ExcelCellFont(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextExcelFontHeight(data)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 4 ->
          new ExcelCellFont(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextExcelColor(data)),
              Optional.empty(),
              Optional.empty());
      case 5 ->
          new ExcelCellFont(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(Boolean.TRUE),
              Optional.empty());
      default ->
          new ExcelCellFont(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(Boolean.FALSE));
    };
  }

  static Optional<CellFontInput> toCellFontInput(@Nullable ExcelCellFont font) {
    if (font == null) {
      return Optional.empty();
    }
    return Optional.of(
        new CellFontInput(
            font.bold(),
            font.italic(),
            font.fontName(),
            font.fontHeight().map(height -> new FontHeightInput.Twips(height.twips())),
            font.fontColor().flatMap(FuzzStyleDecoders::toColorInput),
            font.underline(),
            font.strikeout()));
  }

  private static ExcelHorizontalAlignment nextHorizontalAlignment(GridGrindFuzzData data) {
    ExcelHorizontalAlignment[] values = ExcelHorizontalAlignment.values();
    return values[data.consumeInt(0, values.length - 1)];
  }

  private static ExcelCellAlignment nextAlignment(GridGrindFuzzData data, boolean includeDepth) {
    return new ExcelCellAlignment(
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

  private static ExcelCellFont nextFont(GridGrindFuzzData data, boolean includeName) {
    return new ExcelCellFont(
        data.consumeBoolean() ? Optional.of(Boolean.TRUE) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(Boolean.FALSE) : Optional.empty(),
        includeName
            ? Optional.of(data.consumeBoolean() ? "Aptos" : "Aptos Display")
            : Optional.empty(),
        data.consumeBoolean() ? Optional.of(nextExcelFontHeight(data)) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(nextExcelColor(data)) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(Boolean.TRUE) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(Boolean.FALSE) : Optional.empty());
  }

  private static ExcelCellFill nextFill(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 4)) {
      case 0 -> ExcelCellFill.patternForeground(ExcelFillPattern.SOLID, nextExcelColor(data));
      case 1 ->
          data.consumeBoolean()
              ? ExcelCellFill.patternColors(
                  nextPatternFill(data), nextExcelColor(data), nextExcelColor(data))
              : ExcelCellFill.patternForeground(nextPatternFill(data), nextExcelColor(data));
      case 2 -> ExcelCellFill.patternBackground(nextPatternFill(data), nextExcelColor(data));
      case 3 -> ExcelCellFill.gradient(nextGradientFill(data));
      default ->
          ExcelCellFill.pattern(
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

  private static ExcelCellProtection nextProtection(GridGrindFuzzData data) {
    return new ExcelCellProtection(
        data.consumeBoolean() ? Optional.of(data.consumeBoolean()) : Optional.empty(),
        data.consumeBoolean() ? Optional.of(data.consumeBoolean()) : Optional.empty());
  }

  private static ExcelVerticalAlignment nextVerticalAlignment(GridGrindFuzzData data) {
    ExcelVerticalAlignment[] values = ExcelVerticalAlignment.values();
    return values[data.consumeInt(0, values.length - 1)];
  }

  private static ExcelFontHeight nextExcelFontHeight(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 2)) {
      case 0 -> new ExcelFontHeight(data.consumeInt(20, 640));
      case 1 -> ExcelFontHeight.fromPoints(nextPointHeight(data));
      default -> new ExcelFontHeight(220);
    };
  }

  private static BigDecimal nextPointHeight(GridGrindFuzzData data) {
    return new ExcelFontHeight(data.consumeInt(20, 640)).points();
  }

  private static String nextRgbHex(GridGrindFuzzData data) {
    return "#%02X%02X%02X"
        .formatted(data.consumeInt(0, 255), data.consumeInt(0, 255), data.consumeInt(0, 255));
  }

  private static ExcelColor nextExcelColor(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 2)) {
      case 0 ->
          ExcelColor.rgb(
              nextRgbHex(data),
              data.consumeBoolean() ? Optional.of(nextTint(data)) : Optional.empty());
      case 1 ->
          ExcelColor.theme(
              data.consumeInt(0, 9),
              data.consumeBoolean() ? Optional.of(nextTint(data)) : Optional.empty());
      default ->
          ExcelColor.indexed(
              data.consumeInt(0, 64),
              data.consumeBoolean() ? Optional.of(nextTint(data)) : Optional.empty());
    };
  }

  private static ExcelGradientFill nextGradientFill(GridGrindFuzzData data) {
    boolean linear = data.consumeBoolean();
    List<ExcelGradientStop> stops =
        List.of(
            new ExcelGradientStop(0.0d, nextExcelColor(data)),
            new ExcelGradientStop(1.0d, nextExcelColor(data)));
    if (linear) {
      return ExcelGradientFill.linear(
          data.consumeBoolean()
              ? Optional.of(data.consumeRegularDouble(0.0d, 180.0d))
              : Optional.empty(),
          stops);
    }
    return ExcelGradientFill.path(
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

  private static Optional<ColorInput> toColorInput(ExcelColor color) {
    if (color == null) {
      return Optional.empty();
    }
    return Optional.of(
        switch (color) {
          case ExcelColor.Rgb rgb ->
              rgb.tint().isPresent()
                  ? ColorInput.rgb(rgb.rgb(), rgb.tint().orElseThrow())
                  : ColorInput.rgb(rgb.rgb());
          case ExcelColor.Theme theme ->
              theme.tint().isPresent()
                  ? ColorInput.theme(theme.theme(), theme.tint().orElseThrow())
                  : ColorInput.theme(theme.theme());
          case ExcelColor.Indexed indexed ->
              indexed.tint().isPresent()
                  ? ColorInput.indexed(indexed.indexed(), indexed.tint().orElseThrow())
                  : ColorInput.indexed(indexed.indexed());
        });
  }

  private static ExcelBorder nextExcelBorder(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 4)) {
      case 0 ->
          new ExcelBorder(
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      case 1 ->
          new ExcelBorder(
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.empty());
      case 2 ->
          new ExcelBorder(
              Optional.empty(),
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.empty(),
              Optional.of(nextBorderSide(data)));
      case 3 ->
          new ExcelBorder(
              Optional.of(nextBorderSide(data)),
              Optional.of(nextBorderSide(data)),
              Optional.of(nextBorderSide(data)),
              Optional.of(nextBorderSide(data)),
              Optional.of(nextBorderSide(data)));
      default ->
          new ExcelBorder(
              Optional.empty(),
              Optional.of(nextBorderSide(data)),
              Optional.empty(),
              Optional.of(nextBorderSide(data)),
              Optional.empty());
    };
  }

  private static ExcelBorderSide nextBorderSide(GridGrindFuzzData data) {
    ExcelBorderStyle[] values = ExcelBorderStyle.values();
    ExcelBorderStyle style = values[data.consumeInt(0, values.length - 1)];
    return new ExcelBorderSide(
        Optional.of(style),
        style == ExcelBorderStyle.NONE || !data.consumeBoolean()
            ? Optional.empty()
            : Optional.of(nextExcelColor(data)));
  }
}
