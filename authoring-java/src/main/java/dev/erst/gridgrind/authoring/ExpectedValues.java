package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.dto.CellScalarValue;
import java.util.Objects;

/** Typed authored expected cell-value helpers for assertion-focused Java DSL calls. */
public final class ExpectedValues {
  /** Authored expected effective cell values for assertion helpers. */
  public sealed interface Value
      permits BlankValue, TextValue, NumberValue, BooleanValue, ErrorValue {}

  /** Expected blank effective cell value. */
  public record BlankValue() implements Value {}

  /** Expected text effective cell value. */
  public record TextValue(String text) implements Value {
    public TextValue {
      Objects.requireNonNull(text, "text must not be null");
    }
  }

  /** Expected numeric effective cell value. */
  public record NumberValue(double value) implements Value {}

  /** Expected boolean effective cell value. */
  public record BooleanValue(boolean value) implements Value {}

  /** Expected Excel error effective cell value. */
  public record ErrorValue(String error) implements Value {
    public ErrorValue {
      Objects.requireNonNull(error, "error must not be null");
    }
  }

  private ExpectedValues() {}

  /** Returns one expected blank effective cell value. */
  public static Value blank() {
    return new BlankValue();
  }

  /** Returns one expected text effective cell value. */
  public static Value text(String text) {
    return new TextValue(text);
  }

  /** Returns one expected numeric effective cell value. */
  public static Value number(double number) {
    return new NumberValue(number);
  }

  /** Returns one expected boolean effective cell value. */
  public static Value bool(boolean value) {
    return new BooleanValue(value);
  }

  /** Returns one expected error effective cell value. */
  public static Value error(String error) {
    return new ErrorValue(error);
  }

  static CellScalarValue toCellScalarValue(Value expectedValue) {
    return switch (Objects.requireNonNull(expectedValue, "expectedValue must not be null")) {
      case BlankValue _ -> new CellScalarValue.Blank();
      case TextValue text -> new CellScalarValue.Text(text.text());
      case NumberValue number -> new CellScalarValue.NumberValue(number.value());
      case BooleanValue booleanValue -> new CellScalarValue.BooleanValue(booleanValue.value());
      case ErrorValue error -> new CellScalarValue.ErrorValue(error.error());
    };
  }
}
