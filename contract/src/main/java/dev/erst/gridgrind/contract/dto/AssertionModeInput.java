package dev.erst.gridgrind.contract.dto;

/**
 * Controls whether assertion failures stop execution immediately or complete a terminal check
 * phase.
 */
public enum AssertionModeInput {
  /** Stop at the first failed assertion and return its canonical failure. */
  FAIL_FAST,

  /** Run every terminal-phase assertion and return every assertion outcome. */
  COLLECT;

  /** Returns the default assertion execution mode. */
  public static AssertionModeInput defaults() {
    return FAIL_FAST;
  }

  /** Returns whether this mode is the default fail-fast behavior. */
  public boolean isDefault() {
    return this == FAIL_FAST;
  }

  /** Custom Jackson inclusion filter that omits the default fail-fast mode. */
  public static final class DefaultFilter {
    @Override
    public boolean equals(Object other) {
      return other == null || other == FAIL_FAST;
    }

    @Override
    public int hashCode() {
      return DefaultFilter.class.hashCode();
    }
  }
}
