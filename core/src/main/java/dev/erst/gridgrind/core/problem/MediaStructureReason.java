package dev.erst.gridgrind.core.problem;

/** Closed explanations for established authored media structure failures. */
public enum MediaStructureReason {
  HEADER("required header does not match"),
  TRUNCATED("required data is truncated"),
  LENGTH("a declared length is inconsistent"),
  COUNT("a declared count is inconsistent"),
  OFFSET("a declared offset is invalid for its data"),
  CODE("a coded value is invalid"),
  TERMINATOR("a required terminator is missing"),
  SEQUENCE("records are out of sequence"),
  DIMENSIONS("declared dimensions are invalid"),
  CHECKSUM("an integrity check does not match"),
  EXCESS_DATA("unexpected data follows the structure"),
  METADATA("required metadata is inconsistent");
  private final String description;

  MediaStructureReason(String description) {
    this.description = description;
  }

  /** Returns the complete core-owned explanation for this reason. */
  public String description() {
    return description;
  }
}
