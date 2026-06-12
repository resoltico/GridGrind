package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Canonical scalar cell-value model shared by authored inputs and exact-value assertions. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellScalarValue.Blank.class, name = "BLANK"),
  @JsonSubTypes.Type(value = CellScalarValue.Text.class, name = "TEXT"),
  @JsonSubTypes.Type(value = CellScalarValue.NumberValue.class, name = "NUMBER"),
  @JsonSubTypes.Type(value = CellScalarValue.BooleanValue.class, name = "BOOLEAN"),
  @JsonSubTypes.Type(value = CellScalarValue.ErrorValue.class, name = "ERROR")
})
public sealed interface CellScalarValue
    permits CellScalarValue.Blank,
        CellScalarValue.Text,
        CellScalarValue.NumberValue,
        CellScalarValue.BooleanValue,
        CellScalarValue.ErrorValue {

  /** Expected blank effective cell value. */
  record Blank() implements CellScalarValue {}

  /** Expected exact text effective cell value. */
  record Text(String text) implements CellScalarValue {
    public Text {
      Objects.requireNonNull(text, "text must not be null");
    }
  }

  /** Expected exact numeric effective cell value. */
  record NumberValue(double number) implements CellScalarValue {
    public NumberValue {
      CellInput.Validation.requireFinite(number, "number");
    }
  }

  /** Expected exact boolean effective cell value. */
  record BooleanValue(boolean bool) implements CellScalarValue {}

  /** Expected exact Excel error effective cell value. */
  record ErrorValue(String error) implements CellScalarValue {
    public ErrorValue {
      CellInput.Validation.requireNonBlank(error, "error");
      CellInput.Validation.requireErrorLiteral(error, "error");
    }
  }
}
