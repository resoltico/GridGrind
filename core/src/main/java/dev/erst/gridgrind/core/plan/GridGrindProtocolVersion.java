package dev.erst.gridgrind.core.plan;

/** Version marker for the external GridGrind request/response contract. */
public enum GridGrindProtocolVersion {
  V3;

  /** Returns the protocol version used by this build of GridGrind. */
  public static GridGrindProtocolVersion current() {
    return V3;
  }
}
