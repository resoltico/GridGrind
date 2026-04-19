package dev.erst.gridgrind.contract.dto;

/** Version marker for the external GridGrind request/response contract. */
public enum GridGrindProtocolVersion {
  V1;

  /** Returns the protocol version used by this build of GridGrind. */
  public static GridGrindProtocolVersion current() {
    return V1;
  }
}
