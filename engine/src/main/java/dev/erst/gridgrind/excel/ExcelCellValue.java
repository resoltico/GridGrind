package dev.erst.gridgrind.excel;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Objects;

/** Typed cell values understood by the workbook core. */
public sealed interface ExcelCellValue
    permits ExcelCellValue.BlankValue,
        ExcelCellValue.TextValue,
        ExcelCellValue.NumberValue,
        ExcelCellValue.BooleanValue,
        ExcelCellValue.DateValue,
        ExcelCellValue.DateTimeValue,
        ExcelCellValue.FormulaValue {

  /** Creates an explicit blank cell value. */
  static ExcelCellValue blank() {
    return new BlankValue();
  }

  /** Creates a text cell value. */
  static ExcelCellValue text(String value) {
    return new TextValue(value);
  }

  /** Creates a numeric cell value. */
  static ExcelCellValue number(double value) {
    return new NumberValue(value);
  }

  /** Creates a boolean cell value. */
  static ExcelCellValue bool(boolean value) {
    return new BooleanValue(value);
  }

  /** Creates a date cell value using the workbook date style. */
  static ExcelCellValue date(LocalDate value) {
    return new DateValue(value);
  }

  /** Creates a date-time cell value using the workbook date-time style. */
  static ExcelCellValue dateTime(LocalDateTime value) {
    return new DateTimeValue(value);
  }

  /** Creates a formula cell value. */
  static ExcelCellValue formula(String expression) {
    return new FormulaValue(expression);
  }

  record BlankValue() implements ExcelCellValue {}

  record TextValue(String value) implements ExcelCellValue {
    public TextValue {
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  record NumberValue(double value) implements ExcelCellValue {}

  record BooleanValue(boolean value) implements ExcelCellValue {}

  record DateValue(LocalDate value) implements ExcelCellValue {
    public DateValue {
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  record DateTimeValue(LocalDateTime value) implements ExcelCellValue {
    public DateTimeValue {
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  record FormulaValue(String expression) implements ExcelCellValue {
    public FormulaValue {
      Objects.requireNonNull(expression, "expression must not be null");
      if (expression.isBlank()) {
        throw new IllegalArgumentException("expression must not be blank");
      }
    }
  }
}
