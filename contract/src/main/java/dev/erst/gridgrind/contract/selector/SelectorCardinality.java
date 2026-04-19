package dev.erst.gridgrind.contract.selector;

/** Declares how many targets a selector is allowed to resolve to. */
public enum SelectorCardinality {
  EXACTLY_ONE,
  ZERO_OR_ONE,
  ONE_OR_MORE,
  ANY_NUMBER;

  /** Returns whether the selector may legally resolve to more than one target. */
  public boolean allowsMany() {
    return this == ONE_OR_MORE || this == ANY_NUMBER;
  }

  /** Returns whether the selector may legally resolve to zero targets. */
  public boolean allowsZero() {
    return this == ZERO_OR_ONE || this == ANY_NUMBER;
  }
}
