package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.value.CellGridSpec;
import dev.erst.gridgrind.core.model.value.CellRowSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelRichText;
import dev.erst.gridgrind.excel.ExcelRichTextRun;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Owns bounded cell-value and matrix decoding for Jazzer inputs. */
final class FuzzValueDecoders {
  private FuzzValueDecoders() {}

  static CellValueSpec nextCellInput(GridGrindFuzzData data) {
    Objects.requireNonNull(data, "data must not be null");

    return switch (data.consumeInt(0, 7)) {
      case 0 -> new CellValueSpec.Blank();
      case 1 -> new CellValueSpec.Text(TextSourceSpec.inline(FuzzAddressDecoders.nextText(data)));
      case 2 -> nextRichTextInput(data);
      case 3 -> new CellValueSpec.NumberValue(data.consumeRegularDouble(-1000.0d, 1000.0d));
      case 4 -> new CellValueSpec.BooleanValue(data.consumeBoolean());
      case 5 ->
          new CellValueSpec.Date(
              LocalDate.of(2026, data.consumeInt(1, 12), data.consumeInt(1, 28)));
      case 6 ->
          new CellValueSpec.DateTime(
              LocalDateTime.of(
                  2026,
                  data.consumeInt(1, 12),
                  data.consumeInt(1, 28),
                  data.consumeInt(0, 23),
                  data.consumeInt(0, 59),
                  data.consumeInt(0, 59)));
      default ->
          new CellValueSpec.Formula(TextSourceSpec.inline(FuzzAddressDecoders.nextFormula(data)));
    };
  }

  static ExcelCellValue nextExcelCellValue(GridGrindFuzzData data) {
    Objects.requireNonNull(data, "data must not be null");

    return switch (data.consumeInt(0, 7)) {
      case 0 -> ExcelCellValue.blank();
      case 1 -> ExcelCellValue.text(FuzzAddressDecoders.nextText(data));
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
      default -> ExcelCellValue.formula(FuzzAddressDecoders.nextFormula(data));
    };
  }

  static CellGridSpec nextProtocolGrid(GridGrindFuzzData data, int rowCount, int columnCount) {
    Objects.requireNonNull(data, "data must not be null");

    List<List<CellValueSpec>> rows = new ArrayList<>(rowCount);
    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
      List<CellValueSpec> row = new ArrayList<>(columnCount);
      for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
        row.add(nextCellInput(data));
      }
      rows.add(List.copyOf(row));
    }
    return new CellGridSpec.Typed(List.copyOf(rows));
  }

  static CellRowSpec nextProtocolRow(GridGrindFuzzData data, int columnCount) {
    Objects.requireNonNull(data, "data must not be null");

    List<CellValueSpec> values = new ArrayList<>(columnCount);
    for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      values.add(nextCellInput(data));
    }
    return new CellRowSpec.Typed(List.copyOf(values));
  }

  static List<List<ExcelCellValue>> nextExcelMatrix(
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

  private static CellValueSpec.RichText nextRichTextInput(GridGrindFuzzData data) {
    int runCount = data.consumeInt(1, 3);
    List<RichTextRunSpec> runs = new ArrayList<>(runCount);
    for (int runIndex = 0; runIndex < runCount; runIndex++) {
      var fontPatch = data.consumeBoolean() ? FuzzStyleDecoders.nextRichTextFontPatch(data) : null;
      runs.add(
          new RichTextRunSpec(
              TextSourceSpec.inline(FuzzAddressDecoders.nextText(data)),
              FuzzStyleDecoders.toCellFontInput(fontPatch)));
    }
    return new CellValueSpec.RichText(List.copyOf(runs));
  }

  private static ExcelRichText nextRichText(GridGrindFuzzData data) {
    int runCount = data.consumeInt(1, 3);
    List<ExcelRichTextRun> runs = new ArrayList<>(runCount);
    for (int runIndex = 0; runIndex < runCount; runIndex++) {
      runs.add(
          new ExcelRichTextRun(
              FuzzAddressDecoders.nextText(data),
              data.consumeBoolean()
                  ? Optional.of(FuzzStyleDecoders.nextRichTextFontPatch(data))
                  : Optional.empty()));
    }
    return new ExcelRichText(List.copyOf(runs));
  }
}
