package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Grid-oriented authored cell input with compact homogeneous encodings for range writes. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellGridInput.Typed.class, name = "TYPED"),
  @JsonSubTypes.Type(value = CellGridInput.TextRows.class, name = "TEXT"),
  @JsonSubTypes.Type(value = CellGridInput.NumberRows.class, name = "NUMBER"),
  @JsonSubTypes.Type(value = CellGridInput.BooleanRows.class, name = "BOOLEAN"),
  @JsonSubTypes.Type(value = CellGridInput.ErrorRows.class, name = "ERROR"),
  @JsonSubTypes.Type(value = CellGridInput.DateRows.class, name = "DATE"),
  @JsonSubTypes.Type(value = CellGridInput.DateTimeRows.class, name = "DATE_TIME"),
  @JsonSubTypes.Type(value = CellGridInput.FormulaRows.class, name = "FORMULA")
})
public sealed interface CellGridInput
    permits CellGridInput.Typed,
        CellGridInput.TextRows,
        CellGridInput.NumberRows,
        CellGridInput.BooleanRows,
        CellGridInput.ErrorRows,
        CellGridInput.DateRows,
        CellGridInput.DateTimeRows,
        CellGridInput.FormulaRows {

  /** Returns this authored grid in canonical per-cell form. */
  List<List<CellInput>> toCellInputRows();

  record Typed(List<List<CellInput>> rows) implements CellGridInput {
    public Typed {
      rows = copyTypedRows(rows);
    }

    @Override
    public List<List<CellInput>> toCellInputRows() {
      return rows;
    }
  }

  record TextRows(List<List<String>> rows) implements CellGridInput {
    public TextRows {
      rows = copyScalarRows(rows, "rows");
    }

    @Override
    public List<List<CellInput>> toCellInputRows() {
      return rows.stream()
          .map(
              row ->
                  row.stream()
                      .map(TextSourceInput::inline)
                      .map(CellInput.Text::new)
                      .map(CellInput.class::cast)
                      .toList())
          .toList();
    }
  }

  record NumberRows(List<List<Double>> rows) implements CellGridInput {
    public NumberRows {
      rows = copyScalarRows(rows, "rows");
      rows.forEach(
          row -> row.forEach(value -> CellInput.Validation.requireFinite(value, "rows element")));
    }

    @Override
    public List<List<CellInput>> toCellInputRows() {
      return rows.stream()
          .map(
              row ->
                  row.stream().map(CellInput.NumberValue::new).map(CellInput.class::cast).toList())
          .toList();
    }
  }

  record BooleanRows(List<List<Boolean>> rows) implements CellGridInput {
    public BooleanRows {
      rows = copyScalarRows(rows, "rows");
    }

    @Override
    public List<List<CellInput>> toCellInputRows() {
      return rows.stream()
          .map(
              row ->
                  row.stream().map(CellInput.BooleanValue::new).map(CellInput.class::cast).toList())
          .toList();
    }
  }

  record ErrorRows(List<List<String>> rows) implements CellGridInput {
    public ErrorRows {
      rows = copyScalarRows(rows, "rows");
      rows.forEach(
          row ->
              row.forEach(
                  value -> CellInput.Validation.requireErrorLiteral(value, "rows element")));
    }

    @Override
    public List<List<CellInput>> toCellInputRows() {
      return rows.stream()
          .map(
              row ->
                  row.stream().map(CellInput.ErrorValue::new).map(CellInput.class::cast).toList())
          .toList();
    }
  }

  record DateRows(List<List<LocalDate>> rows) implements CellGridInput {
    public DateRows {
      rows = copyScalarRows(rows, "rows");
    }

    @Override
    public List<List<CellInput>> toCellInputRows() {
      return rows.stream()
          .map(row -> row.stream().map(CellInput.Date::new).map(CellInput.class::cast).toList())
          .toList();
    }
  }

  record DateTimeRows(List<List<LocalDateTime>> rows) implements CellGridInput {
    public DateTimeRows {
      rows = copyScalarRows(rows, "rows");
    }

    @Override
    public List<List<CellInput>> toCellInputRows() {
      return rows.stream()
          .map(row -> row.stream().map(CellInput.DateTime::new).map(CellInput.class::cast).toList())
          .toList();
    }
  }

  record FormulaRows(List<List<String>> rows) implements CellGridInput {
    public FormulaRows {
      rows =
          copyScalarRows(rows, "rows").stream()
              .map(
                  row ->
                      row.stream()
                          .map(
                              value ->
                                  CellInput.Validation.normalizeInlineFormula(
                                      value, "rows element"))
                          .toList())
              .toList();
    }

    @Override
    public List<List<CellInput>> toCellInputRows() {
      return rows.stream()
          .map(
              row ->
                  row.stream()
                      .map(TextSourceInput::inline)
                      .map(CellInput.Formula::new)
                      .map(CellInput.class::cast)
                      .toList())
          .toList();
    }
  }

  private static List<List<CellInput>> copyTypedRows(List<List<CellInput>> rows) {
    Objects.requireNonNull(rows, "rows must not be null");
    List<List<CellInput>> copy = new ArrayList<>(rows.size());
    if (rows.isEmpty()) {
      throw new IllegalArgumentException("rows must not be empty");
    }
    int expectedWidth = -1;
    for (List<CellInput> row : rows) {
      Objects.requireNonNull(row, "rows must not contain null rows");
      if (row.isEmpty()) {
        throw new IllegalArgumentException("rows must not contain empty rows");
      }
      if (expectedWidth < 0) {
        expectedWidth = row.size();
      } else if (row.size() != expectedWidth) {
        throw new IllegalArgumentException("rows must describe a rectangular matrix");
      }
      copy.add(copyTypedRow(row));
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

  private static List<CellInput> copyTypedRow(List<CellInput> row) {
    List<CellInput> rowCopy = new ArrayList<>(row.size());
    for (CellInput value : row) {
      rowCopy.add(Objects.requireNonNull(value, "rows must not contain null cell values"));
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
