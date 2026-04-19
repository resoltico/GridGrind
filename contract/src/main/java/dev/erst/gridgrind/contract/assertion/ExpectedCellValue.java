package dev.erst.gridgrind.contract.assertion;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Expected effective cell value used by exact-value assertions. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ExpectedCellValue.Blank.class, name = "BLANK"),
  @JsonSubTypes.Type(value = ExpectedCellValue.Text.class, name = "TEXT"),
  @JsonSubTypes.Type(value = ExpectedCellValue.NumericValue.class, name = "NUMBER"),
  @JsonSubTypes.Type(value = ExpectedCellValue.BooleanValue.class, name = "BOOLEAN"),
  @JsonSubTypes.Type(value = ExpectedCellValue.ErrorValue.class, name = "ERROR")
})
public sealed interface ExpectedCellValue
    permits ExpectedCellValue.Blank,
        ExpectedCellValue.Text,
        ExpectedCellValue.NumericValue,
        ExpectedCellValue.BooleanValue,
        ExpectedCellValue.ErrorValue {

  /** Expects the effective cell value to be blank. */
  record Blank() implements ExpectedCellValue {}

  /** Expects the effective cell value to be an exact string. */
  record Text(String text) implements ExpectedCellValue {
    public Text {
      Objects.requireNonNull(text, "text must not be null");
    }
  }

  /** Expects the effective cell value to be an exact finite number. */
  record NumericValue(Double number) implements ExpectedCellValue {
    public NumericValue {
      number = AssertionSupport.requireFiniteNumber(number, "number");
    }
  }

  /** Expects the effective cell value to be a boolean. */
  record BooleanValue(Boolean value) implements ExpectedCellValue {
    public BooleanValue {
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  /** Expects the effective cell value to be an exact Excel error string such as #REF!. */
  record ErrorValue(String error) implements ExpectedCellValue {
    public ErrorValue {
      error = AssertionSupport.requireNonBlank(error, "error");
    }
  }
}
