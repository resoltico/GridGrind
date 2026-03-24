package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelCellValue;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Objects;

/** JSON-friendly typed cell input used at the agent protocol boundary. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellInput.Blank.class, name = "BLANK"),
  @JsonSubTypes.Type(value = CellInput.Text.class, name = "TEXT"),
  @JsonSubTypes.Type(value = CellInput.Number.class, name = "NUMBER"),
  @JsonSubTypes.Type(value = CellInput.BooleanValue.class, name = "BOOLEAN"),
  @JsonSubTypes.Type(value = CellInput.Date.class, name = "DATE"),
  @JsonSubTypes.Type(value = CellInput.DateTime.class, name = "DATE_TIME"),
  @JsonSubTypes.Type(value = CellInput.Formula.class, name = "FORMULA")
})
public sealed interface CellInput {

  /** Converts the protocol value to the workbook-core representation. */
  ExcelCellValue toExcelCellValue();

  /** Blank (empty) cell input that clears the target cell. */
  record Blank() implements CellInput {
    @Override
    public ExcelCellValue toExcelCellValue() {
      return ExcelCellValue.blank();
    }
  }

  /** Plain string cell input. */
  record Text(String text) implements CellInput {
    public Text {
      Validation.required(text, "text");
    }

    @Override
    public ExcelCellValue toExcelCellValue() {
      return ExcelCellValue.text(text);
    }
  }

  /** Numeric cell input stored as a double. */
  record Number(Double number) implements CellInput {
    public Number {
      Validation.required(number, "number");
    }

    @Override
    public ExcelCellValue toExcelCellValue() {
      return ExcelCellValue.number(number);
    }
  }

  /** Boolean cell input. */
  record BooleanValue(Boolean bool) implements CellInput {
    public BooleanValue {
      Validation.required(bool, "bool");
    }

    @Override
    public ExcelCellValue toExcelCellValue() {
      return ExcelCellValue.bool(bool);
    }
  }

  /** Excel formula cell input. */
  record Formula(String formula) implements CellInput {
    public Formula {
      Validation.required(formula, "formula");
    }

    @Override
    public ExcelCellValue toExcelCellValue() {
      return ExcelCellValue.formula(formula);
    }
  }

  /** ISO-8601 date cell input formatted as a date value in Excel. */
  record Date(String date) implements CellInput {
    public Date {
      LocalDate.parse(Validation.required(date, "date"));
    }

    @Override
    public ExcelCellValue toExcelCellValue() {
      return ExcelCellValue.date(LocalDate.parse(date));
    }
  }

  /** ISO-8601 date-time cell input formatted as a date-time value in Excel. */
  record DateTime(String dateTime) implements CellInput {
    public DateTime {
      LocalDateTime.parse(Validation.required(dateTime, "dateTime"));
    }

    @Override
    public ExcelCellValue toExcelCellValue() {
      return ExcelCellValue.dateTime(LocalDateTime.parse(dateTime));
    }
  }

  /** Null-checking helpers for CellInput compact constructors. */
  final class Validation {
    private Validation() {}

    static <T> T required(T value, String fieldName) {
      if (value == null) {
        throw new IllegalArgumentException(fieldName + " must not be null");
      }
      return value;
    }
  }
}
