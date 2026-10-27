package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** Tagged machine-readable validation constraint published for one scalar catalog field. */
@Union(discriminator = "type")
public sealed interface FieldConstraint
    permits FieldConstraint.NonBlank,
        FieldConstraint.StringPattern,
        FieldConstraint.LengthRange,
        FieldConstraint.NumberRange,
        FieldConstraint.Integral,
        FieldConstraint.PathSuffix {
  /** Requires a nonblank string. */
  @Variant(id = "NON_BLANK", summary = "Requires a nonblank string.")
  record NonBlank() implements FieldConstraint {}

  /** Requires a string to match one regular-expression pattern. */
  @Variant(
      id = "STRING_PATTERN",
      summary = "Requires a string to match one regular-expression pattern.")
  record StringPattern(String pattern) implements FieldConstraint {
    public StringPattern {
      pattern = TextValidation.requireNonBlank(pattern, "pattern");
    }
  }

  /** Requires a string length inside one inclusive range. */
  @Variant(id = "LENGTH_RANGE", summary = "Requires a string length inside one inclusive range.")
  record LengthRange(int minimum, int maximum) implements FieldConstraint {
    public LengthRange {
      if (minimum < 0 || maximum < minimum) {
        throw new IllegalArgumentException("length range must be nonnegative and ordered");
      }
    }
  }

  /** Requires a number inside one inclusive finite range. */
  @Variant(id = "NUMBER_RANGE", summary = "Requires a number inside one inclusive finite range.")
  record NumberRange(double minimum, double maximum) implements FieldConstraint {
    public NumberRange {
      if (!Double.isFinite(minimum) || !Double.isFinite(maximum) || maximum < minimum) {
        throw new IllegalArgumentException("number range must be finite and ordered");
      }
    }
  }

  /** Requires a JSON number to be integral. */
  @Variant(id = "INTEGRAL", summary = "Requires a JSON number to be integral.")
  record Integral() implements FieldConstraint {}

  /** Requires one path to end in a literal suffix. */
  @Variant(id = "PATH_SUFFIX", summary = "Requires one path to end in a literal suffix.")
  record PathSuffix(String suffix) implements FieldConstraint {
    public PathSuffix {
      suffix = TextValidation.requireNonBlank(suffix, "suffix");
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
