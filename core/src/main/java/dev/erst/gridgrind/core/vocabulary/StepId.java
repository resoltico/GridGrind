package dev.erst.gridgrind.core.vocabulary;

/** A workbook step identifier using the published identifier alphabet. */
public record StepId(String value) implements WireString {
  private static final java.util.regex.Pattern PATTERN =
      java.util.regex.Pattern.compile(ProtocolConstraintValues.STEP_ID_PATTERN);

  public StepId {
    java.util.Objects.requireNonNull(value, "stepId must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException("stepId must not be blank");
    }
    if (!PATTERN.matcher(value).matches()) {
      throw new IllegalArgumentException("stepId must match [A-Za-z0-9._-]+");
    }
  }

  @Override
  public String toString() {
    return value;
  }
}
