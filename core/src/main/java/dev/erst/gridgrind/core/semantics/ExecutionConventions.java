package dev.erst.gridgrind.core.semantics;

/** Fixed shared execution/rendering semantics; these are not host environment defaults. */
public enum ExecutionConventions {
  STANDARD;

  /** Returns the fixed rendering language tag, independent of host defaults. */
  public String localeTag() {
    return "en-US";
  }

  /** Returns the neutral offset used only when bridging civil values to instant-based APIs. */
  public java.time.ZoneOffset civilDateBridgeOffset() {
    return java.time.ZoneOffset.UTC;
  }
}
