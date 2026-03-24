package dev.erst.gridgrind.protocol;

/** Version marker for the external GridGrind request/response contract. */
public enum GridGrindProtocolVersion {
  V1;

  public static GridGrindProtocolVersion current() {
    return V1;
  }
}
