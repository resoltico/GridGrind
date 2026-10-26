package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.Objects;

/** Canonical scalar effective cell-value model used by exact-value assertions. */
@Union(discriminator = "type", catalog = true, displayOrder = 0)
public sealed interface CellScalarValue
    permits dev.erst.gridgrind.core.model.value.CellScalarValue.Blank,
        dev.erst.gridgrind.core.model.value.CellScalarValue.Text,
        dev.erst.gridgrind.core.model.value.CellScalarValue.NumberValue,
        dev.erst.gridgrind.core.model.value.CellScalarValue.BooleanValue,
        dev.erst.gridgrind.core.model.value.CellScalarValue.ErrorValue {

  /** Expected blank effective cell value. */
  @Variant(id = "BLANK", summary = "Require the effective cell value to be blank.")
  record Blank() implements CellScalarValue {}

  /** Expected exact text effective cell value. */
  @Variant(id = "TEXT", summary = "Require the effective cell value to be one exact string.")
  record Text(String text) implements CellScalarValue {
    public Text {
      Objects.requireNonNull(text, "text must not be null");
    }
  }

  /** Expected exact numeric effective cell value. */
  @Variant(
      id = "NUMBER",
      summary = "Require the effective cell value to be one exact finite number.")
  record NumberValue(double number) implements CellScalarValue {
    public NumberValue {
      CellInput.Validation.requireFinite(number, "number");
    }
  }

  /** Expected exact boolean effective cell value. */
  @Variant(id = "BOOLEAN", summary = "Require the effective cell value to be true or false.")
  record BooleanValue(boolean bool) implements CellScalarValue {}

  /** Expected exact GridGrind-owned error literal as the effective cell value. */
  @Variant(
      id = "ERROR",
      summary =
          "Require the effective cell value to be one exact GridGrind-owned error literal, including canonical Excel tokens and evaluation-only states such as #CIRCULAR_REF!.")
  record ErrorValue(String error) implements CellScalarValue {
    public ErrorValue {
      error = CellInput.Validation.requireNonBlank(error, "error");
      error = CellErrorLiteralValidation.requireReportedErrorLiteral(error, "error");
    }
  }
}
