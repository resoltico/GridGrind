package dev.erst.gridgrind.core.vocabulary;

import dev.erst.gridgrind.core.limits.Limit;

/** A workbook step identifier using the published identifier alphabet. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_STEP_ID_TOKEN_SHAPE
})
public record StepId(String value) implements WireString {
  private static final java.util.regex.Pattern PATTERN =
      java.util.regex.Pattern.compile(Limit.STEP_ID_PATTERN);

  public StepId {
    value = TextValidation.requireNonBlank(value, "stepId");
    if (!PATTERN.matcher(value).matches()) {
      throw new IllegalArgumentException("stepId must match [A-Za-z0-9._-]+");
    }
  }

  @Override
  public String toString() {
    return value;
  }
}
