package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Grid-oriented authored cell input with compact homogeneous encodings for range writes. */
@Union(discriminator = "type", catalog = true, displayOrder = 18)
public sealed interface CellGridSpec
    permits dev.erst.gridgrind.core.model.value.CellGridSpec.Typed,
        dev.erst.gridgrind.core.model.value.CellGridSpec.TextRows,
        dev.erst.gridgrind.core.model.value.CellGridSpec.NumberRows,
        dev.erst.gridgrind.core.model.value.CellGridSpec.BooleanRows,
        dev.erst.gridgrind.core.model.value.CellGridSpec.ErrorRows,
        dev.erst.gridgrind.core.model.value.CellGridSpec.DateRows,
        dev.erst.gridgrind.core.model.value.CellGridSpec.DateTimeRows,
        dev.erst.gridgrind.core.model.value.CellGridSpec.FormulaRows {

  /** Returns this authored grid in canonical per-cell form. */
  List<List<CellValueSpec>> toCellInputRows();

  @Variant(id = "TYPED", summary = "Write one rectangular grid expressed as per-cell typed rows.")
  record Typed(List<List<CellValueSpec>> cells) implements CellGridSpec {
    public Typed {
      cells = copyTypedRows(cells, "cells");
    }

    @Override
    public List<List<CellValueSpec>> toCellInputRows() {
      return cells;
    }
  }

  @Variant(id = "TEXT", summary = "Write one rectangular grid using homogeneous string rows.")
  record TextRows(List<List<String>> cells) implements CellGridSpec {
    public TextRows {
      cells = copyScalarRows(cells, "cells");
    }

    @Override
    public List<List<CellValueSpec>> toCellInputRows() {
      return cells.stream()
          .map(
              row ->
                  row.stream()
                      .map(TextSourceSpec::inline)
                      .map(CellValueSpec.Text::new)
                      .map(CellValueSpec.class::cast)
                      .toList())
          .toList();
    }
  }

  @Variant(id = "NUMBER", summary = "Write one rectangular grid using homogeneous numeric rows.")
  record NumberRows(List<List<Double>> cells) implements CellGridSpec {
    public NumberRows {
      cells = copyScalarRows(cells, "cells");
      cells.forEach(
          row ->
              row.forEach(value -> CellValueSpec.Validation.requireFinite(value, "cells element")));
    }

    @Override
    public List<List<CellValueSpec>> toCellInputRows() {
      return cells.stream()
          .map(
              row ->
                  row.stream()
                      .map(CellValueSpec.NumberValue::new)
                      .map(CellValueSpec.class::cast)
                      .toList())
          .toList();
    }
  }

  @Variant(id = "BOOLEAN", summary = "Write one rectangular grid using homogeneous boolean rows.")
  record BooleanRows(List<List<Boolean>> cells) implements CellGridSpec {
    public BooleanRows {
      cells = copyScalarRows(cells, "cells");
    }

    @Override
    public List<List<CellValueSpec>> toCellInputRows() {
      return cells.stream()
          .map(
              row ->
                  row.stream()
                      .map(CellValueSpec.BooleanValue::new)
                      .map(CellValueSpec.class::cast)
                      .toList())
          .toList();
    }
  }

  @Variant(
      id = "ERROR",
      summary = "Write one rectangular grid using homogeneous stored Excel error literal rows.")
  record ErrorRows(List<List<String>> cells) implements CellGridSpec {
    public ErrorRows {
      cells = copyScalarRows(cells, "cells");
      cells.forEach(
          row ->
              row.forEach(
                  value ->
                      CellErrorLiteralValidation.requireStoredErrorLiteral(
                          value, "cells element")));
    }

    @Override
    public List<List<CellValueSpec>> toCellInputRows() {
      return cells.stream()
          .map(
              row ->
                  row.stream()
                      .map(CellValueSpec.ErrorValue::new)
                      .map(CellValueSpec.class::cast)
                      .toList())
          .toList();
    }
  }

  @Variant(id = "DATE", summary = "Write one rectangular grid using homogeneous date rows.")
  record DateRows(List<List<LocalDate>> cells) implements CellGridSpec {
    public DateRows {
      cells = copyScalarRows(cells, "cells");
    }

    @Override
    public List<List<CellValueSpec>> toCellInputRows() {
      return cells.stream()
          .map(
              row ->
                  row.stream().map(CellValueSpec.Date::new).map(CellValueSpec.class::cast).toList())
          .toList();
    }
  }

  @Variant(
      id = "DATE_TIME",
      summary = "Write one rectangular grid using homogeneous date-time rows.")
  record DateTimeRows(List<List<LocalDateTime>> cells) implements CellGridSpec {
    public DateTimeRows {
      cells = copyScalarRows(cells, "cells");
    }

    @Override
    public List<List<CellValueSpec>> toCellInputRows() {
      return cells.stream()
          .map(
              row ->
                  row.stream()
                      .map(CellValueSpec.DateTime::new)
                      .map(CellValueSpec.class::cast)
                      .toList())
          .toList();
    }
  }

  @Variant(
      id = "FORMULA",
      summary = "Write one rectangular grid using homogeneous formula text rows.")
  record FormulaRows(List<List<String>> cells) implements CellGridSpec {
    public FormulaRows {
      cells =
          copyScalarRows(cells, "cells").stream()
              .map(
                  row ->
                      row.stream()
                          .map(
                              value ->
                                  CellValueSpec.Validation.normalizeInlineFormula(
                                      value, "cells element"))
                          .toList())
              .toList();
    }

    @Override
    public List<List<CellValueSpec>> toCellInputRows() {
      return cells.stream()
          .map(
              row ->
                  row.stream()
                      .map(TextSourceSpec::inline)
                      .map(CellValueSpec.Formula::new)
                      .map(CellValueSpec.class::cast)
                      .toList())
          .toList();
    }
  }

  private static List<List<CellValueSpec>> copyTypedRows(
      List<List<CellValueSpec>> rows, String fieldName) {
    Objects.requireNonNull(rows, fieldName + " must not be null");
    List<List<CellValueSpec>> copy = new ArrayList<>(rows.size());
    if (rows.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    int expectedWidth = -1;
    for (List<CellValueSpec> row : rows) {
      Objects.requireNonNull(row, fieldName + " must not contain null rows");
      if (row.isEmpty()) {
        throw new IllegalArgumentException(fieldName + " must not contain empty rows");
      }
      if (expectedWidth < 0) {
        expectedWidth = row.size();
      } else if (row.size() != expectedWidth) {
        throw new IllegalArgumentException(fieldName + " must describe a rectangular matrix");
      }
      copy.add(copyTypedRow(row, fieldName));
    }
    return List.copyOf(copy);
  }

  private static <T> List<List<T>> copyScalarRows(List<List<T>> rows, String fieldName) {
    Objects.requireNonNull(rows, fieldName + " must not be null");
    if (rows.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    List<List<T>> copy = new ArrayList<>(rows.size());
    int expectedWidth = -1;
    for (List<T> row : rows) {
      Objects.requireNonNull(row, fieldName + " must not contain null rows");
      if (row.isEmpty()) {
        throw new IllegalArgumentException(fieldName + " must not contain empty rows");
      }
      if (expectedWidth < 0) {
        expectedWidth = row.size();
      } else if (row.size() != expectedWidth) {
        throw new IllegalArgumentException(fieldName + " must describe a rectangular matrix");
      }
      copy.add(copyScalarRow(row, fieldName));
    }
    return List.copyOf(copy);
  }

  private static List<CellValueSpec> copyTypedRow(List<CellValueSpec> row, String fieldName) {
    List<CellValueSpec> rowCopy = new ArrayList<>(row.size());
    for (CellValueSpec value : row) {
      rowCopy.add(Objects.requireNonNull(value, fieldName + " must not contain null cell values"));
    }
    return List.copyOf(rowCopy);
  }

  private static <T> List<T> copyScalarRow(List<T> row, String fieldName) {
    List<T> rowCopy = new ArrayList<>(row.size());
    for (T value : row) {
      rowCopy.add(Objects.requireNonNull(value, fieldName + " must not contain null values"));
    }
    return List.copyOf(rowCopy);
  }
}
