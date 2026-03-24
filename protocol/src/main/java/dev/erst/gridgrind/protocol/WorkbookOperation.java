package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;
import java.util.Objects;

/** One validated workbook operation expressed in protocol form. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = WorkbookOperation.EnsureSheet.class, name = "ENSURE_SHEET"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetCell.class, name = "SET_CELL"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetRange.class, name = "SET_RANGE"),
  @JsonSubTypes.Type(value = WorkbookOperation.ClearRange.class, name = "CLEAR_RANGE"),
  @JsonSubTypes.Type(value = WorkbookOperation.ApplyStyle.class, name = "APPLY_STYLE"),
  @JsonSubTypes.Type(value = WorkbookOperation.AppendRow.class, name = "APPEND_ROW"),
  @JsonSubTypes.Type(value = WorkbookOperation.AutoSizeColumns.class, name = "AUTO_SIZE_COLUMNS"),
  @JsonSubTypes.Type(value = WorkbookOperation.EvaluateFormulas.class, name = "EVALUATE_FORMULAS"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.ForceFormulaRecalculationOnOpen.class,
      name = "FORCE_FORMULA_RECALCULATION_ON_OPEN")
})
public sealed interface WorkbookOperation {

  /** Ensures a sheet with the given name exists, creating it if absent. */
  record EnsureSheet(String sheetName) implements WorkbookOperation {
    public EnsureSheet {
      Validation.requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Sets a single cell to the given value. */
  record SetCell(String sheetName, String address, CellInput value) implements WorkbookOperation {
    public SetCell {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(address, "address");
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  /** Sets a rectangular region of cells from a row-major grid of values. */
  record SetRange(String sheetName, String range, List<List<CellInput>> rows)
      implements WorkbookOperation {
    public SetRange {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
      rows = Validation.copyRows(rows);
      Validation.requireRectangularRows(rows);
    }
  }

  /** Clears all cell values and styles within the specified range. */
  record ClearRange(String sheetName, String range) implements WorkbookOperation {
    public ClearRange {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
    }
  }

  /** Applies a style patch to every cell in the specified range. */
  record ApplyStyle(String sheetName, String range, CellStyleInput style)
      implements WorkbookOperation {
    public ApplyStyle {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
      Objects.requireNonNull(style, "style must not be null");
    }
  }

  /** Appends a new row of values after the last occupied row on the sheet. */
  record AppendRow(String sheetName, List<CellInput> values) implements WorkbookOperation {
    public AppendRow {
      Validation.requireNonBlank(sheetName, "sheetName");
      values = values == null ? List.of() : List.copyOf(values);
      if (values.isEmpty()) {
        throw new IllegalArgumentException("values must not be empty");
      }
      for (CellInput item : values) {
        Objects.requireNonNull(item, "values must not contain nulls");
      }
    }
  }

  /** Auto-sizes the specified columns to fit their content. */
  record AutoSizeColumns(String sheetName, List<String> columns) implements WorkbookOperation {
    public AutoSizeColumns {
      Validation.requireNonBlank(sheetName, "sheetName");
      columns = columns == null ? List.of() : List.copyOf(columns);
      if (columns.isEmpty()) {
        throw new IllegalArgumentException("columns must not be empty");
      }
      for (String column : columns) {
        Validation.requireNonBlank(column, "column");
      }
    }
  }

  /** Evaluates all formulas in the workbook at operation time. */
  record EvaluateFormulas() implements WorkbookOperation {}

  /** Marks the workbook so that Excel recalculates all formulas on next open. */
  record ForceFormulaRecalculationOnOpen() implements WorkbookOperation {}

  /** Returns the SCREAMING_SNAKE_CASE type name of this operation as used in the wire protocol. */
  default String operationType() {
    return switch (this) {
      case EnsureSheet _ -> "ENSURE_SHEET";
      case SetCell _ -> "SET_CELL";
      case SetRange _ -> "SET_RANGE";
      case ClearRange _ -> "CLEAR_RANGE";
      case ApplyStyle _ -> "APPLY_STYLE";
      case AppendRow _ -> "APPEND_ROW";
      case AutoSizeColumns _ -> "AUTO_SIZE_COLUMNS";
      case EvaluateFormulas _ -> "EVALUATE_FORMULAS";
      case ForceFormulaRecalculationOnOpen _ -> "FORCE_FORMULA_RECALCULATION_ON_OPEN";
    };
  }

  /**
   * Returns the sheet name if this operation targets a specific sheet, or null for operations that
   * do not target a sheet (EvaluateFormulas, ForceFormulaRecalculationOnOpen).
   *
   * <p>Null is permitted here because this is a protocol-layer default method on a sealed interface
   * used for wire serialization; internal code must use a switch expression instead.
   */
  default String extractSheetName() {
    return switch (this) {
      case EnsureSheet op -> op.sheetName();
      case SetCell op -> op.sheetName();
      case SetRange op -> op.sheetName();
      case ClearRange op -> op.sheetName();
      case ApplyStyle op -> op.sheetName();
      case AppendRow op -> op.sheetName();
      case AutoSizeColumns op -> op.sheetName();
      case EvaluateFormulas _, ForceFormulaRecalculationOnOpen _ -> null;
    };
  }

  /**
   * Returns the cell address if this operation targets a single cell, or null for all other
   * subtypes.
   *
   * <p>Null is permitted here because this is a protocol-layer default method on a sealed interface
   * used for wire serialization; internal code must use a switch expression instead.
   */
  default String extractAddress() {
    return switch (this) {
      case SetCell op -> op.address();
      case EnsureSheet _, SetRange _, ClearRange _, ApplyStyle _,
          AppendRow _, AutoSizeColumns _, EvaluateFormulas _,
          ForceFormulaRecalculationOnOpen _ -> null;
    };
  }

  /**
   * Returns the range if this operation targets a range (SetRange, ClearRange, ApplyStyle), or
   * null for all other subtypes.
   *
   * <p>Null is permitted here because this is a protocol-layer default method on a sealed interface
   * used for wire serialization; internal code must use a switch expression instead.
   */
  default String extractRange() {
    return switch (this) {
      case SetRange op -> op.range();
      case ClearRange op -> op.range();
      case ApplyStyle op -> op.range();
      case EnsureSheet _, SetCell _, AppendRow _, AutoSizeColumns _,
          EvaluateFormulas _, ForceFormulaRecalculationOnOpen _ -> null;
    };
  }

  /**
   * Returns the cell input value if this is a SetCell operation, or null for all other subtypes.
   *
   * <p>Null is permitted here because this is a protocol-layer default method on a sealed interface
   * used for wire serialization; internal code must use a switch expression instead.
   */
  default CellInput extractValue() {
    return switch (this) {
      case SetCell op -> op.value();
      case EnsureSheet _, SetRange _, ClearRange _, ApplyStyle _,
          AppendRow _, AutoSizeColumns _, EvaluateFormulas _,
          ForceFormulaRecalculationOnOpen _ -> null;
    };
  }

  /** Shared validation helpers for WorkbookOperation compact constructors. */
  final class Validation {
    private Validation() {}

    static void requireNonBlank(String value, String fieldName) {
      Objects.requireNonNull(value, fieldName + " must not be null");
      if (value.isBlank()) {
        throw new IllegalArgumentException(fieldName + " must not be blank");
      }
    }

    static List<List<CellInput>> copyRows(List<List<CellInput>> rows) {
      if (rows == null) {
        return List.of();
      }
      return rows.stream().map(row -> row == null ? null : List.copyOf(row)).toList();
    }

    static void requireRectangularRows(List<List<CellInput>> rows) {
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
        for (CellInput value : row) {
          Objects.requireNonNull(value, "rows must not contain null cell values");
        }
      }
    }
  }
}
