package dev.erst.gridgrind.core.vocabulary;

/** A nonblank caller-owned workbook plan identifier. */
public record PlanId(String value) implements WireString {
  public PlanId {
    value = TextValidation.requireNonBlank(value, "planId");
  }

  @Override
  public String toString() {
    return value;
  }
}
