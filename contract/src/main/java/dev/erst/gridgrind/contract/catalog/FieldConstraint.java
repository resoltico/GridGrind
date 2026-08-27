package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Tagged machine-readable validation constraint published for one scalar catalog field. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = FieldConstraint.NonBlank.class, name = "NON_BLANK"),
  @JsonSubTypes.Type(value = FieldConstraint.StringPattern.class, name = "STRING_PATTERN"),
  @JsonSubTypes.Type(value = FieldConstraint.LengthRange.class, name = "LENGTH_RANGE"),
  @JsonSubTypes.Type(value = FieldConstraint.NumberRange.class, name = "NUMBER_RANGE"),
  @JsonSubTypes.Type(value = FieldConstraint.Integral.class, name = "INTEGRAL"),
  @JsonSubTypes.Type(value = FieldConstraint.PathSuffix.class, name = "PATH_SUFFIX")
})
public sealed interface FieldConstraint
    permits FieldConstraint.NonBlank,
        FieldConstraint.StringPattern,
        FieldConstraint.LengthRange,
        FieldConstraint.NumberRange,
        FieldConstraint.Integral,
        FieldConstraint.PathSuffix {
  /** Requires a nonblank string. */
  record NonBlank() implements FieldConstraint {}

  /** Requires a string to match one regular-expression pattern. */
  record StringPattern(String pattern) implements FieldConstraint {
    public StringPattern {
      pattern = CatalogRecordValidation.requireNonBlank(pattern, "pattern");
    }
  }

  /** Requires a string length inside one inclusive range. */
  record LengthRange(int minimum, int maximum) implements FieldConstraint {
    public LengthRange {
      if (minimum < 0 || maximum < minimum) {
        throw new IllegalArgumentException("length range must be nonnegative and ordered");
      }
    }
  }

  /** Requires a number inside one inclusive finite range. */
  record NumberRange(double minimum, double maximum) implements FieldConstraint {
    public NumberRange {
      if (!Double.isFinite(minimum) || !Double.isFinite(maximum) || maximum < minimum) {
        throw new IllegalArgumentException("number range must be finite and ordered");
      }
    }
  }

  /** Requires a JSON number to be integral. */
  record Integral() implements FieldConstraint {}

  /** Requires one path to end in a literal suffix. */
  record PathSuffix(String suffix) implements FieldConstraint {
    public PathSuffix {
      suffix = CatalogRecordValidation.requireNonBlank(suffix, "suffix");
      if (!suffix.startsWith(".")) {
        throw new IllegalArgumentException("suffix must begin with .");
      }
    }
  }

  /** Returns the stable type discriminator used for canonical constraint ordering. */
  default String type() {
    return switch (this) {
      case NonBlank _ -> "NON_BLANK";
      case StringPattern _ -> "STRING_PATTERN";
      case LengthRange _ -> "LENGTH_RANGE";
      case NumberRange _ -> "NUMBER_RANGE";
      case Integral _ -> "INTEGRAL";
      case PathSuffix _ -> "PATH_SUFFIX";
    };
  }

  /** Returns a deterministic lexical key after the stable constraint type. */
  default String sortKey() {
    return switch (this) {
      case NonBlank _ -> "";
      case StringPattern stringPattern -> stringPattern.pattern();
      case LengthRange lengthRange -> lengthRange.minimum() + ":" + lengthRange.maximum();
      case NumberRange numberRange -> numberRange.minimum() + ":" + numberRange.maximum();
      case Integral _ -> "";
      case PathSuffix pathSuffix -> pathSuffix.suffix();
    };
  }
}
