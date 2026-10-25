package dev.erst.gridgrind.core.vocabulary;

/** A nonblank caller-owned workbook plan identifier. */
public record PlanId(String value) implements WireString {
  public PlanId {
    java.util.Objects.requireNonNull(value, "planId must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException("planId must not be blank");
    }
  }

  @Override
  public String toString() {
    return value;
  }
}
