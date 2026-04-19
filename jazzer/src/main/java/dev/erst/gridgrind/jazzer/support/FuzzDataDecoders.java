package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelCellAlignment;
import dev.erst.gridgrind.excel.ExcelCellFill;
import dev.erst.gridgrind.excel.ExcelCellFont;
import dev.erst.gridgrind.excel.ExcelCellProtection;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelGradientFill;
import dev.erst.gridgrind.excel.ExcelGradientStop;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelRichText;
import dev.erst.gridgrind.excel.ExcelRichTextRun;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Decodes bounded structured values from a Jazzer data provider. */
public final class FuzzDataDecoders {
  private static final List<String> VALID_COLUMNS = List.of("A", "B", "C", "D", "E", "F", "G", "H");

  private FuzzDataDecoders() {}

  /** Returns a bounded sheet name, optionally valid for Excel sheet creation. */
  public static String nextSheetName(GridGrindFuzzData data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");

    if (!valid) {
      return switch (data.consumeInt(0, 5)) {
        case 0 -> "";
        case 1 -> " ";
        case 2 -> "Bad/Name";
        case 3 -> "Bad*Name";
        case 4 -> "Bad[Name]";
        default -> "ABCDEFGHIJKLMNOPQRSTUVWXYZABCDEF"; // 32 chars, exceeds the 31-char limit
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
  public static String nextCellAddress(GridGrindFuzzData data, boolean valid) {
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
  public static String nextNonBlankCellAddress(GridGrindFuzzData data, boolean valid) {
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
  public static String nextRange(GridGrindFuzzData data, boolean valid) {
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

  /**
   * Returns a bounded non-blank rectangular A1-style range that may still be semantically invalid.
   */
  public static String nextNonBlankRange(GridGrindFuzzData data, boolean valid) {
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
  public static CellInput nextCellInput(GridGrindFuzzData data) {
    Objects.requireNonNull(data, "data must not be null");

    return switch (data.consumeInt(0, 7)) {
      case 0 -> new CellInput.Blank();
      case 1 -> new CellInput.Text(TextSourceInput.inline(nextText(data)));
      case 2 -> nextRichTextInput(data);
      case 3 -> new CellInput.Numeric(data.consumeRegularDouble(-1000.0d, 1000.0d));
      case 4 -> new CellInput.BooleanValue(data.consumeBoolean());
      case 5 ->
          new CellInput.Date(LocalDate.of(2026, data.consumeInt(1, 12), data.consumeInt(1, 28)));
      case 6 ->
          new CellInput.DateTime(
              LocalDateTime.of(
                  2026,
                  data.consumeInt(1, 12),
                  data.consumeInt(1, 28),
                  data.consumeInt(0, 23),
                  data.consumeInt(0, 59),
                  data.consumeInt(0, 59)));
      default -> new CellInput.Formula(TextSourceInput.inline(nextFormula(data)));
    };
  }

  /** Returns a workbook-core cell value. */
  public static ExcelCellValue nextExcelCellValue(GridGrindFuzzData data) {
    Objects.requireNonNull(data, "data must not be null");

    return switch (data.consumeInt(0, 7)) {
      case 0 -> ExcelCellValue.blank();
      case 1 -> ExcelCellValue.text(nextText(data));
      case 2 -> ExcelCellValue.richText(nextRichText(data));
      case 3 -> ExcelCellValue.number(data.consumeRegularDouble(-1000.0d, 1000.0d));
      case 4 -> ExcelCellValue.bool(data.consumeBoolean());
      case 5 ->
          ExcelCellValue.date(LocalDate.of(2026, data.consumeInt(1, 12), data.consumeInt(1, 28)));
      case 6 ->
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

  private static CellInput.RichText nextRichTextInput(GridGrindFuzzData data) {
    int runCount = data.consumeInt(1, 3);
    List<RichTextRunInput> runs = new ArrayList<>(runCount);
    for (int runIndex = 0; runIndex < runCount; runIndex++) {
      ExcelCellFont fontPatch = data.consumeBoolean() ? nextRichTextFontPatch(data) : null;
      runs.add(
          new RichTextRunInput(TextSourceInput.inline(nextText(data)), toCellFontInput(fontPatch)));
    }
    return new CellInput.RichText(List.copyOf(runs));
  }

  private static ExcelRichText nextRichText(GridGrindFuzzData data) {
    int runCount = data.consumeInt(1, 3);
    List<ExcelRichTextRun> runs = new ArrayList<>(runCount);
    for (int runIndex = 0; runIndex < runCount; runIndex++) {
      runs.add(
          new ExcelRichTextRun(
              nextText(data), data.consumeBoolean() ? nextRichTextFontPatch(data) : null));
    }
    return new ExcelRichText(List.copyOf(runs));
  }

  /** Returns a bounded style patch. */
  public static ExcelCellStyle nextStyle(GridGrindFuzzData data) {
    Objects.requireNonNull(data, "data must not be null");

    return switch (data.consumeInt(0, 7)) {
      case 0 ->
          new ExcelCellStyle(
              data.consumeBoolean() ? "0.00" : "yyyy-mm-dd", null, null, null, null, null);
      case 1 -> new ExcelCellStyle(null, nextAlignment(data, true), null, null, null, null);
      case 2 -> new ExcelCellStyle(null, null, nextFont(data, true), null, null, null);
      case 3 ->
          new ExcelCellStyle(
              null,
              nextAlignment(data, true),
              nextFont(data, true),
              nextFill(data),
              nextExcelBorder(data),
              nextProtection(data));
      case 4 ->
          new ExcelCellStyle(
              null, null, nextFont(data, false), nextFill(data), nextExcelBorder(data), null);
      case 5 -> new ExcelCellStyle(null, null, null, nextFill(data), null, nextProtection(data));
      case 6 ->
          new ExcelCellStyle(
              null, nextAlignment(data, true), null, null, nextExcelBorder(data), null);
      default ->
          new ExcelCellStyle(null, null, nextFont(data, false), null, null, nextProtection(data));
    };
  }

  /** Returns a non-empty rectangular matrix of protocol cell inputs. */
  public static List<List<CellInput>> nextProtocolMatrix(
      GridGrindFuzzData data, int rowCount, int columnCount) {
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
      GridGrindFuzzData data, int rowCount, int columnCount) {
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

  private static ExcelHorizontalAlignment nextHorizontalAlignment(GridGrindFuzzData data) {
    ExcelHorizontalAlignment[] values = ExcelHorizontalAlignment.values();
    return values[data.consumeInt(0, values.length - 1)];
  }

  private static ExcelCellAlignment nextAlignment(GridGrindFuzzData data, boolean includeDepth) {
    return new ExcelCellAlignment(
        data.consumeBoolean() ? Boolean.TRUE : null,
        nextHorizontalAlignment(data),
        nextVerticalAlignment(data),
        includeDepth && data.consumeBoolean() ? data.consumeInt(0, 180) : null,
        includeDepth && data.consumeBoolean() ? data.consumeInt(0, 8) : null);
  }

  private static ExcelCellFont nextFont(GridGrindFuzzData data, boolean includeName) {
    return new ExcelCellFont(
        data.consumeBoolean() ? Boolean.TRUE : null,
        data.consumeBoolean() ? Boolean.FALSE : null,
        includeName ? (data.consumeBoolean() ? "Aptos" : "Aptos Display") : null,
        data.consumeBoolean() ? nextExcelFontHeight(data) : null,
        data.consumeBoolean() ? nextExcelColor(data) : null,
        data.consumeBoolean() ? Boolean.TRUE : null,
        data.consumeBoolean() ? Boolean.FALSE : null);
  }

  private static ExcelCellFont nextRichTextFontPatch(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 6)) {
      case 0 -> new ExcelCellFont(Boolean.TRUE, null, null, null, null, null, null);
      case 1 -> new ExcelCellFont(null, Boolean.FALSE, null, null, null, null, null);
      case 2 -> new ExcelCellFont(null, null, "Aptos", null, null, null, null);
      case 3 -> new ExcelCellFont(null, null, null, nextExcelFontHeight(data), null, null, null);
      case 4 -> new ExcelCellFont(null, null, null, null, nextExcelColor(data), null, null);
      case 5 -> new ExcelCellFont(null, null, null, null, null, Boolean.TRUE, null);
      default -> new ExcelCellFont(null, null, null, null, null, null, Boolean.FALSE);
    };
  }

  private static ExcelCellFill nextFill(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 4)) {
      case 0 -> new ExcelCellFill(ExcelFillPattern.SOLID, nextExcelColor(data), null, null);
      case 1 ->
          new ExcelCellFill(
              nextPatternFill(data),
              nextExcelColor(data),
              data.consumeBoolean() ? nextExcelColor(data) : null,
              null);
      case 2 -> new ExcelCellFill(null, nextExcelColor(data), null, null);
      case 3 -> new ExcelCellFill(null, null, null, nextGradientFill(data));
      default -> new ExcelCellFill(nextPatternFill(data), null, null, null);
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
        data.consumeBoolean() ? data.consumeBoolean() : null,
        data.consumeBoolean() ? data.consumeBoolean() : null);
  }

  private static ExcelVerticalAlignment nextVerticalAlignment(GridGrindFuzzData data) {
    ExcelVerticalAlignment[] values = ExcelVerticalAlignment.values();
    return values[data.consumeInt(0, values.length - 1)];
  }

  private static String nextText(GridGrindFuzzData data) {
    int length = data.consumeInt(1, 16);
    StringBuilder builder = new StringBuilder(length);
    for (int index = 0; index < length; index++) {
      builder.append((char) data.consumeInt('A', 'z'));
    }
    String result = builder.toString().trim();
    return result.isBlank() ? "X" : result;
  }

  private static String nextFormula(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 5)) {
      case 0 -> "SUM(A1:A2)";
      case 1 -> "A1+A2";
      case 2 -> "1/0";
      case 3 -> "IF(A1>0,1,0)";
      case 4 -> "BAD(";
      default -> nextText(data);
    };
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
      case 0 -> new ExcelColor(nextRgbHex(data));
      case 1 ->
          new ExcelColor(
              null, data.consumeInt(0, 9), null, data.consumeBoolean() ? nextTint(data) : null);
      default ->
          new ExcelColor(
              null, null, data.consumeInt(0, 64), data.consumeBoolean() ? nextTint(data) : null);
    };
  }

  private static ExcelGradientFill nextGradientFill(GridGrindFuzzData data) {
    boolean linear = data.consumeBoolean();
    return new ExcelGradientFill(
        linear ? "LINEAR" : "PATH",
        linear && data.consumeBoolean() ? data.consumeRegularDouble(0.0d, 180.0d) : null,
        linear || !data.consumeBoolean() ? null : data.consumeRegularDouble(0.0d, 1.0d),
        linear || !data.consumeBoolean() ? null : data.consumeRegularDouble(0.0d, 1.0d),
        linear || !data.consumeBoolean() ? null : data.consumeRegularDouble(0.0d, 1.0d),
        linear || !data.consumeBoolean() ? null : data.consumeRegularDouble(0.0d, 1.0d),
        List.of(
            new ExcelGradientStop(0.0d, nextExcelColor(data)),
            new ExcelGradientStop(1.0d, nextExcelColor(data))));
  }

  private static double nextTint(GridGrindFuzzData data) {
    return data.consumeRegularDouble(-1.0d, 1.0d);
  }

  private static CellFontInput toCellFontInput(ExcelCellFont font) {
    if (font == null) {
      return null;
    }
    return new CellFontInput(
        font.bold(),
        font.italic(),
        font.fontName(),
        font.fontHeight() == null ? null : new FontHeightInput.Twips(font.fontHeight().twips()),
        font.fontColor() == null ? null : font.fontColor().rgb(),
        font.fontColor() == null ? null : font.fontColor().theme(),
        font.fontColor() == null ? null : font.fontColor().indexed(),
        font.fontColor() == null ? null : font.fontColor().tint(),
        font.underline(),
        font.strikeout());
  }

  private static ExcelBorder nextExcelBorder(GridGrindFuzzData data) {
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

  private static ExcelBorderSide nextBorderSide(GridGrindFuzzData data) {
    ExcelBorderStyle[] values = ExcelBorderStyle.values();
    ExcelBorderStyle style = values[data.consumeInt(0, values.length - 1)];
    return new ExcelBorderSide(
        style,
        style == ExcelBorderStyle.NONE || !data.consumeBoolean() ? null : nextExcelColor(data));
  }
}
