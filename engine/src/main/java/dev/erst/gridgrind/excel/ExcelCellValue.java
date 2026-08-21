package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.contract.dto.FormulaTextValidation;
import dev.erst.gridgrind.excel.foundation.ExcelStoredCellErrorLiteral;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Objects;

/** Typed cell values understood by the workbook core. */
public sealed interface ExcelCellValue
    permits ExcelCellValue.BlankValue,
        ExcelCellValue.TextValue,
        ExcelCellValue.RichTextValue,
        ExcelCellValue.NumberValue,
        ExcelCellValue.BooleanValue,
        ExcelCellValue.ErrorValue,
        ExcelCellValue.DateValue,
        ExcelCellValue.DateTimeValue,
        ExcelCellValue.FormulaValue,
        ExcelCellValue.RawFormulaValue {

  /** Creates an explicit blank cell value. */
  static ExcelCellValue blank() {
    return new BlankValue();
  }

  /** Creates a text cell value. */
  static ExcelCellValue text(String value) {
    return new TextValue(value);
  }

  /** Creates a structured rich-text string cell value. */
  static ExcelCellValue richText(ExcelRichText value) {
    return new RichTextValue(value);
  }

  /** Creates a numeric cell value. */
  static ExcelCellValue number(double value) {
    return new NumberValue(value);
  }

  /** Creates a boolean cell value. */
  static ExcelCellValue bool(boolean value) {
    return new BooleanValue(value);
  }

  /** Creates one stored OOXML error cell value such as {@code #REF!}. */
  static ExcelCellValue error(String value) {
    return new ErrorValue(value);
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

  /** Creates one opaque OOXML formula-body value. */
  static ExcelCellValue rawFormula(String expression) {
    return new RawFormulaValue(expression);
  }

  record BlankValue() implements ExcelCellValue {}

  record TextValue(String value) implements ExcelCellValue {
    public TextValue {
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  record RichTextValue(ExcelRichText value) implements ExcelCellValue {
    public RichTextValue {
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  record NumberValue(double value) implements ExcelCellValue {}

  record BooleanValue(boolean value) implements ExcelCellValue {}

  /** Stored OOXML error cell literal. */
  record ErrorValue(String value) implements ExcelCellValue {
    public ErrorValue {
      Objects.requireNonNull(value, "value must not be null");
      value = ExcelStoredCellErrorLiteral.fromWireValue(value).wireValue();
    }
  }

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
      expression = FormulaTextValidation.requireNormalFormulaBody(expression, "expression");
    }
  }

  /** Formula character data written without sending its body through POI's formula parser. */
  record RawFormulaValue(String expression) implements ExcelCellValue {
    public RawFormulaValue {
      expression = FormulaTextValidation.requireRawFormulaBody(expression, "expression");
    }
  }
}
