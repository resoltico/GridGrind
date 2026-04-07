package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.protocol.dto.CellAlignmentInput;
import dev.erst.gridgrind.protocol.dto.CellBorderInput;
import dev.erst.gridgrind.protocol.dto.CellBorderSideInput;
import dev.erst.gridgrind.protocol.dto.CellFillInput;
import dev.erst.gridgrind.protocol.dto.CellFontInput;
import dev.erst.gridgrind.protocol.dto.CellProtectionInput;
import dev.erst.gridgrind.protocol.dto.CellStyleInput;
import dev.erst.gridgrind.protocol.dto.FontHeightInput;
import java.math.BigDecimal;
import java.util.Objects;

/** Creates bounded protocol style patches for fuzz harnesses. */
public final class WorkbookStyleInputs {
  private WorkbookStyleInputs() {}

  /** Returns a bounded protocol style patch. */
  public static CellStyleInput nextStyleInput(GridGrindFuzzData data) {
    Objects.requireNonNull(data, "data must not be null");

    return switch (data.consumeInt(0, 7)) {
      case 0 ->
          new CellStyleInput(
              data.consumeBoolean() ? "0.00" : "yyyy-mm-dd", null, null, null, null, null);
      case 1 ->
          new CellStyleInput(null, nextAlignmentInput(data, true), null, null, null, null);
      case 2 -> new CellStyleInput(null, null, nextFontInput(data, true), null, null, null);
      case 3 ->
          new CellStyleInput(
              null,
              nextAlignmentInput(data, true),
              nextFontInput(data, true),
              nextFillInput(data),
              nextProtocolBorder(data),
              nextProtectionInput(data));
      case 4 ->
          new CellStyleInput(
              null, null, nextFontInput(data, false), nextFillInput(data), nextProtocolBorder(data), null);
      case 5 -> new CellStyleInput(null, null, null, nextFillInput(data), null, nextProtectionInput(data));
      case 6 ->
          new CellStyleInput(null, nextAlignmentInput(data, true), null, null, nextProtocolBorder(data), null);
      default ->
          new CellStyleInput(null, null, nextFontInput(data, false), null, null, nextProtectionInput(data));
    };
  }

  private static CellAlignmentInput nextAlignmentInput(GridGrindFuzzData data, boolean includeDepth) {
    return new CellAlignmentInput(
        data.consumeBoolean() ? Boolean.TRUE : null,
        nextHorizontalAlignment(data),
        nextVerticalAlignment(data),
        includeDepth && data.consumeBoolean() ? data.consumeInt(0, 180) : null,
        includeDepth && data.consumeBoolean() ? data.consumeInt(0, 8) : null);
  }

  private static CellFontInput nextFontInput(GridGrindFuzzData data, boolean includeName) {
    return new CellFontInput(
        data.consumeBoolean() ? Boolean.TRUE : null,
        data.consumeBoolean() ? Boolean.FALSE : null,
        includeName ? (data.consumeBoolean() ? "Aptos" : "Aptos Display") : null,
        data.consumeBoolean() ? nextFontHeightInput(data) : null,
        data.consumeBoolean() ? nextRgbHex(data) : null,
        data.consumeBoolean() ? Boolean.TRUE : null,
        data.consumeBoolean() ? Boolean.FALSE : null);
  }

  private static CellFillInput nextFillInput(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 3)) {
      case 0 -> new CellFillInput(ExcelFillPattern.SOLID, nextRgbHex(data), null);
      case 1 ->
          new CellFillInput(
              nextPatternFill(data),
              nextRgbHex(data),
              data.consumeBoolean() ? nextRgbHex(data) : null);
      case 2 -> new CellFillInput(null, nextRgbHex(data), null);
      default -> new CellFillInput(nextPatternFill(data), null, null);
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
    return "#%02X%02X%02X".formatted(
        data.consumeInt(0, 255), data.consumeInt(0, 255), data.consumeInt(0, 255));
  }

  private static CellBorderInput nextProtocolBorder(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 4)) {
      case 0 -> new CellBorderInput(nextBorderSide(data), null, null, null, null);
      case 1 -> new CellBorderInput(nextBorderSide(data), null, nextBorderSide(data), null, null);
      case 2 -> new CellBorderInput(null, nextBorderSide(data), null, null, nextBorderSide(data));
      case 3 ->
          new CellBorderInput(
              nextBorderSide(data),
              nextBorderSide(data),
              nextBorderSide(data),
              nextBorderSide(data),
              nextBorderSide(data));
      default -> new CellBorderInput(null, nextBorderSide(data), null, nextBorderSide(data), null);
    };
  }

  private static CellBorderSideInput nextBorderSide(GridGrindFuzzData data) {
    ExcelBorderStyle[] values = ExcelBorderStyle.values();
    ExcelBorderStyle style = values[data.consumeInt(0, values.length - 1)];
    return new CellBorderSideInput(
        style, style == ExcelBorderStyle.NONE || !data.consumeBoolean() ? null : nextRgbHex(data));
  }

  private static CellProtectionInput nextProtectionInput(GridGrindFuzzData data) {
    return new CellProtectionInput(
        data.consumeBoolean() ? data.consumeBoolean() : null,
        data.consumeBoolean() ? data.consumeBoolean() : null);
  }

  private static ExcelHorizontalAlignment nextHorizontalAlignment(GridGrindFuzzData data) {
    ExcelHorizontalAlignment[] values = ExcelHorizontalAlignment.values();
    return values[data.consumeInt(0, values.length - 1)];
  }

  private static ExcelVerticalAlignment nextVerticalAlignment(GridGrindFuzzData data) {
    ExcelVerticalAlignment[] values = ExcelVerticalAlignment.values();
    return values[data.consumeInt(0, values.length - 1)];
  }
}
