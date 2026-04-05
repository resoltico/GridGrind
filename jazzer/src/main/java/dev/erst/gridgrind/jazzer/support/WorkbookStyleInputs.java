package dev.erst.gridgrind.jazzer.support;

import com.code_intelligence.jazzer.api.FuzzedDataProvider;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.protocol.dto.CellBorderInput;
import dev.erst.gridgrind.protocol.dto.CellBorderSideInput;
import dev.erst.gridgrind.protocol.dto.CellStyleInput;
import dev.erst.gridgrind.protocol.dto.FontHeightInput;
import java.math.BigDecimal;
import java.util.Objects;

/** Creates bounded protocol style patches for fuzz harnesses. */
public final class WorkbookStyleInputs {
  private WorkbookStyleInputs() {}

  /** Returns a bounded protocol style patch. */
  public static CellStyleInput nextStyleInput(FuzzedDataProvider data) {
    Objects.requireNonNull(data, "data must not be null");

    ExcelHorizontalAlignment[] horizontalValues = ExcelHorizontalAlignment.values();
    ExcelVerticalAlignment[] verticalValues = ExcelVerticalAlignment.values();
    return switch (data.consumeInt(0, 7)) {
      case 0 ->
          new CellStyleInput(
              data.consumeBoolean() ? "0.00" : "yyyy-mm-dd",
              null,
              null,
              Boolean.TRUE,
              horizontalValues[data.consumeInt(0, horizontalValues.length - 1)],
              verticalValues[data.consumeInt(0, verticalValues.length - 1)],
              null,
              null,
              null,
              null,
              null,
              null,
              null);
      case 1 ->
          new CellStyleInput(
              null,
              data.consumeBoolean() ? Boolean.TRUE : null,
              data.consumeBoolean() ? Boolean.TRUE : null,
              data.consumeBoolean() ? Boolean.TRUE : null,
              null,
              null,
              "Aptos",
              nextFontHeightInput(data),
              nextRgbHex(data),
              data.consumeBoolean() ? Boolean.TRUE : null,
              data.consumeBoolean() ? Boolean.FALSE : null,
              null,
              null);
      case 2 ->
          new CellStyleInput(
              null,
              null,
              null,
              null,
              horizontalValues[data.consumeInt(0, horizontalValues.length - 1)],
              verticalValues[data.consumeInt(0, verticalValues.length - 1)],
              data.consumeBoolean() ? "Aptos Display" : null,
              nextFontHeightInput(data),
              data.consumeBoolean() ? nextRgbHex(data) : null,
              data.consumeBoolean() ? data.consumeBoolean() : null,
              data.consumeBoolean() ? data.consumeBoolean() : null,
              nextRgbHex(data),
              null);
      case 3 ->
          new CellStyleInput(
              null,
              Boolean.TRUE,
              Boolean.FALSE,
              Boolean.TRUE,
              horizontalValues[data.consumeInt(0, horizontalValues.length - 1)],
              verticalValues[data.consumeInt(0, verticalValues.length - 1)],
              "Aptos",
              nextFontHeightInput(data),
              nextRgbHex(data),
              Boolean.TRUE,
              data.consumeBoolean() ? Boolean.TRUE : Boolean.FALSE,
              nextRgbHex(data),
              nextProtocolBorder(data));
      case 4 ->
          new CellStyleInput(
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              nextRgbHex(data),
              null,
              null,
              nextRgbHex(data),
              nextProtocolBorder(data));
      case 5 ->
          new CellStyleInput(
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              nextFontHeightInput(data),
              null,
              null,
              null,
              null,
              null);
      case 6 ->
          new CellStyleInput(
              null,
              null,
              null,
              Boolean.TRUE,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              nextProtocolBorder(data));
      default ->
          new CellStyleInput(
              null,
              null,
              null,
              Boolean.TRUE,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null);
    };
  }

  private static FontHeightInput nextFontHeightInput(FuzzedDataProvider data) {
    return switch (data.consumeInt(0, 2)) {
      case 0 -> new FontHeightInput.Twips(data.consumeInt(20, 640));
      case 1 -> new FontHeightInput.Points(nextPointHeight(data));
      default -> new FontHeightInput.Twips(220);
    };
  }

  private static BigDecimal nextPointHeight(FuzzedDataProvider data) {
    return new ExcelFontHeight(data.consumeInt(20, 640)).points();
  }

  private static String nextRgbHex(FuzzedDataProvider data) {
    return "#%02X%02X%02X".formatted(
        data.consumeInt(0, 255), data.consumeInt(0, 255), data.consumeInt(0, 255));
  }

  private static CellBorderInput nextProtocolBorder(FuzzedDataProvider data) {
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

  private static CellBorderSideInput nextBorderSide(FuzzedDataProvider data) {
    ExcelBorderStyle[] values = ExcelBorderStyle.values();
    return new CellBorderSideInput(values[data.consumeInt(0, values.length - 1)]);
  }
}
