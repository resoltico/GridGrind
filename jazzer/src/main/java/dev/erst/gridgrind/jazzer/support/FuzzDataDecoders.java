package dev.erst.gridgrind.jazzer.support;

import com.code_intelligence.jazzer.api.FuzzedDataProvider;
import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.protocol.CellInput;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Decodes bounded structured values from a Jazzer data provider. */
public final class FuzzDataDecoders {
  private static final List<String> VALID_COLUMNS =
      List.of("A", "B", "C", "D", "E", "F", "G", "H");

  private FuzzDataDecoders() {}

  /** Returns a bounded sheet name, optionally valid for Excel sheet creation. */
  public static String nextSheetName(FuzzedDataProvider data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");

    if (!valid) {
      return switch (data.consumeInt(0, 4)) {
        case 0 -> "";
        case 1 -> " ";
        case 2 -> "Bad/Name";
        case 3 -> "Bad*Name";
        default -> "Bad[Name]";
      };
    }

    StringBuilder builder = new StringBuilder();
    int length = data.consumeInt(1, 12);
    for (int index = 0; index < length; index++) {
      char value = (char) data.consumeInt('A', 'Z');
      if ("[]:*?/\\'".indexOf(value) >= 0) {
        value = 'S';
      }
      builder.append(value);
    }
    return builder.toString();
  }

  /** Returns a bounded A1-style cell address. */
  public static String nextCellAddress(FuzzedDataProvider data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");

    if (!valid) {
      return switch (data.consumeInt(0, 4)) {
        case 0 -> "";
        case 1 -> "ZZZ999999";
        case 2 -> "A0";
        case 3 -> "R1C1";
        default -> "1A";
      };
    }

    String column = VALID_COLUMNS.get(data.consumeInt(0, VALID_COLUMNS.size() - 1));
    int row = data.consumeInt(1, 25);
    return column + row;
  }

  /** Returns a bounded non-blank A1-style cell address that may still be semantically invalid. */
  public static String nextNonBlankCellAddress(FuzzedDataProvider data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");

    if (valid) {
      return nextCellAddress(data, true);
    }

    return switch (data.consumeInt(0, 3)) {
      case 0 -> "ZZZ999999";
      case 1 -> "A0";
      case 2 -> "R1C1";
      default -> "1A";
    };
  }

  /** Returns a bounded rectangular A1-style range. */
  public static String nextRange(FuzzedDataProvider data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");

    if (!valid) {
      return switch (data.consumeInt(0, 4)) {
        case 0 -> "";
        case 1 -> "A1:";
        case 2 -> "A1:1B";
        case 3 -> "A0:B2";
        default -> "B2:A1";
      };
    }

    int firstColumnIndex = data.consumeInt(0, VALID_COLUMNS.size() - 2);
    int lastColumnIndex = data.consumeInt(firstColumnIndex, firstColumnIndex + 1);
    int firstRow = data.consumeInt(1, 6);
    int lastRow = data.consumeInt(firstRow, firstRow + 2);
    return VALID_COLUMNS.get(firstColumnIndex)
        + firstRow
        + ":"
        + VALID_COLUMNS.get(lastColumnIndex)
        + lastRow;
  }

  /** Returns a bounded non-blank rectangular A1-style range that may still be semantically invalid. */
  public static String nextNonBlankRange(FuzzedDataProvider data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");

    if (valid) {
      return nextRange(data, true);
    }

    return switch (data.consumeInt(0, 3)) {
      case 0 -> "A1:";
      case 1 -> "A1:1B";
      case 2 -> "A0:B2";
      default -> "B2:A1";
    };
  }

  /** Returns a protocol-layer cell input value. */
  public static CellInput nextCellInput(FuzzedDataProvider data) {
    Objects.requireNonNull(data, "data must not be null");

    return switch (data.consumeInt(0, 6)) {
      case 0 -> new CellInput.Blank();
      case 1 -> new CellInput.Text(nextText(data));
      case 2 -> new CellInput.Numeric(data.consumeRegularDouble(-1000.0d, 1000.0d));
      case 3 -> new CellInput.BooleanValue(data.consumeBoolean());
      case 4 -> new CellInput.Date(LocalDate.of(2026, data.consumeInt(1, 12), data.consumeInt(1, 28)));
      case 5 ->
          new CellInput.DateTime(
              LocalDateTime.of(
                  2026,
                  data.consumeInt(1, 12),
                  data.consumeInt(1, 28),
                  data.consumeInt(0, 23),
                  data.consumeInt(0, 59),
                  data.consumeInt(0, 59)));
      default -> new CellInput.Formula(nextFormula(data));
    };
  }

  /** Returns a workbook-core cell value. */
  public static ExcelCellValue nextExcelCellValue(FuzzedDataProvider data) {
    Objects.requireNonNull(data, "data must not be null");

    return switch (data.consumeInt(0, 6)) {
      case 0 -> ExcelCellValue.blank();
      case 1 -> ExcelCellValue.text(nextText(data));
      case 2 -> ExcelCellValue.number(data.consumeRegularDouble(-1000.0d, 1000.0d));
      case 3 -> ExcelCellValue.bool(data.consumeBoolean());
      case 4 -> ExcelCellValue.date(LocalDate.of(2026, data.consumeInt(1, 12), data.consumeInt(1, 28)));
      case 5 ->
          ExcelCellValue.dateTime(
              LocalDateTime.of(
                  2026,
                  data.consumeInt(1, 12),
                  data.consumeInt(1, 28),
                  data.consumeInt(0, 23),
                  data.consumeInt(0, 59),
                  data.consumeInt(0, 59)));
      default -> ExcelCellValue.formula(nextFormula(data));
    };
  }

  /** Returns a bounded style patch. */
  public static ExcelCellStyle nextStyle(FuzzedDataProvider data) {
    Objects.requireNonNull(data, "data must not be null");

    return switch (data.consumeInt(0, 7)) {
      case 0 -> ExcelCellStyle.alignment(nextHorizontalAlignment(data), nextVerticalAlignment(data));
      case 1 -> ExcelCellStyle.emphasis(Boolean.TRUE, data.consumeBoolean() ? Boolean.FALSE : null);
      case 2 ->
          new ExcelCellStyle(
              data.consumeBoolean() ? "0.00" : null,
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
      case 3 ->
          new ExcelCellStyle(
              null,
              null,
              null,
              null,
              null,
              null,
              data.consumeBoolean() ? "Aptos" : "Aptos Display",
              nextExcelFontHeight(data),
              data.consumeBoolean() ? nextRgbHex(data) : null,
              data.consumeBoolean() ? data.consumeBoolean() : null,
              data.consumeBoolean() ? data.consumeBoolean() : null,
              null,
              null);
      case 4 ->
          new ExcelCellStyle(
              null,
              data.consumeBoolean() ? data.consumeBoolean() : null,
              data.consumeBoolean() ? data.consumeBoolean() : null,
              data.consumeBoolean() ? data.consumeBoolean() : null,
              nextHorizontalAlignment(data),
              nextVerticalAlignment(data),
              data.consumeBoolean() ? "Aptos" : null,
              nextExcelFontHeight(data),
              nextRgbHex(data),
              data.consumeBoolean() ? data.consumeBoolean() : null,
              data.consumeBoolean() ? data.consumeBoolean() : null,
              nextRgbHex(data),
              null);
      case 5 ->
          new ExcelCellStyle(
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              nextRgbHex(data),
              nextExcelBorder(data));
      case 6 ->
          new ExcelCellStyle(
              data.consumeBoolean() ? "yyyy-mm-dd" : "#,##0.00",
              data.consumeBoolean() ? Boolean.TRUE : null,
              data.consumeBoolean() ? Boolean.TRUE : null,
              Boolean.TRUE,
              nextHorizontalAlignment(data),
              nextVerticalAlignment(data),
              "Aptos",
              nextExcelFontHeight(data),
              nextRgbHex(data),
              Boolean.TRUE,
              data.consumeBoolean() ? Boolean.TRUE : Boolean.FALSE,
              nextRgbHex(data),
              nextExcelBorder(data));
      default ->
          new ExcelCellStyle(
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

  /** Returns a non-empty rectangular matrix of protocol cell inputs. */
  public static List<List<CellInput>> nextProtocolMatrix(
      FuzzedDataProvider data, int rowCount, int columnCount) {
    Objects.requireNonNull(data, "data must not be null");

    List<List<CellInput>> rows = new ArrayList<>(rowCount);
    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
      List<CellInput> row = new ArrayList<>(columnCount);
      for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
        row.add(nextCellInput(data));
      }
      rows.add(List.copyOf(row));
    }
    return List.copyOf(rows);
  }

  /** Returns a non-empty rectangular matrix of workbook-core cell values. */
  public static List<List<ExcelCellValue>> nextExcelMatrix(
      FuzzedDataProvider data, int rowCount, int columnCount) {
    Objects.requireNonNull(data, "data must not be null");

    List<List<ExcelCellValue>> rows = new ArrayList<>(rowCount);
    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
      List<ExcelCellValue> row = new ArrayList<>(columnCount);
      for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
        row.add(nextExcelCellValue(data));
      }
      rows.add(List.copyOf(row));
    }
    return List.copyOf(rows);
  }

  private static ExcelHorizontalAlignment nextHorizontalAlignment(FuzzedDataProvider data) {
    ExcelHorizontalAlignment[] values = ExcelHorizontalAlignment.values();
    return values[data.consumeInt(0, values.length - 1)];
  }

  private static ExcelVerticalAlignment nextVerticalAlignment(FuzzedDataProvider data) {
    ExcelVerticalAlignment[] values = ExcelVerticalAlignment.values();
    return values[data.consumeInt(0, values.length - 1)];
  }

  private static String nextText(FuzzedDataProvider data) {
    int length = data.consumeInt(1, 16);
    StringBuilder builder = new StringBuilder(length);
    for (int index = 0; index < length; index++) {
      builder.append((char) data.consumeInt('A', 'z'));
    }
    String result = builder.toString().trim();
    return result.isBlank() ? "X" : result;
  }

  private static String nextFormula(FuzzedDataProvider data) {
    return switch (data.consumeInt(0, 5)) {
      case 0 -> "SUM(A1:A2)";
      case 1 -> "A1+A2";
      case 2 -> "1/0";
      case 3 -> "IF(A1>0,1,0)";
      case 4 -> "BAD(";
      default -> nextText(data);
    };
  }

  private static ExcelFontHeight nextExcelFontHeight(FuzzedDataProvider data) {
    return switch (data.consumeInt(0, 2)) {
      case 0 -> new ExcelFontHeight(data.consumeInt(20, 640));
      case 1 -> ExcelFontHeight.fromPoints(nextPointHeight(data));
      default -> new ExcelFontHeight(220);
    };
  }

  private static BigDecimal nextPointHeight(FuzzedDataProvider data) {
    return new ExcelFontHeight(data.consumeInt(20, 640)).points();
  }

  private static String nextRgbHex(FuzzedDataProvider data) {
    return "#%02X%02X%02X".formatted(
        data.consumeInt(0, 255), data.consumeInt(0, 255), data.consumeInt(0, 255));
  }

  private static ExcelBorder nextExcelBorder(FuzzedDataProvider data) {
    return switch (data.consumeInt(0, 4)) {
      case 0 -> new ExcelBorder(nextBorderSide(data), null, null, null, null);
      case 1 -> new ExcelBorder(nextBorderSide(data), null, nextBorderSide(data), null, null);
      case 2 -> new ExcelBorder(null, nextBorderSide(data), null, null, nextBorderSide(data));
      case 3 ->
          new ExcelBorder(
              nextBorderSide(data),
              nextBorderSide(data),
              nextBorderSide(data),
              nextBorderSide(data),
              nextBorderSide(data));
      default -> new ExcelBorder(null, nextBorderSide(data), null, nextBorderSide(data), null);
    };
  }

  private static ExcelBorderSide nextBorderSide(FuzzedDataProvider data) {
    ExcelBorderStyle[] values = ExcelBorderStyle.values();
    return new ExcelBorderSide(values[data.consumeInt(0, values.length - 1)]);
  }
}
