package dev.erst.gridgrind.core.vocabulary;

import dev.erst.gridgrind.core.metadata.FieldRule;

/** The single supported integer version of the external document contract. */
public record ProtocolVersion(
    @FieldRule(
            kind = FieldRule.Kind.NUMBER_RANGE,
            minimum = ProtocolVersion.NUMBER,
            maximum = ProtocolVersion.NUMBER)
        int value)
    implements WireScalar {
  /** Authoritative wire version for the foundations contract. */
  public static final int NUMBER = 1;

  /** The protocol version emitted by this build. */
  public static final ProtocolVersion CURRENT = new ProtocolVersion(NUMBER);

  public ProtocolVersion {
    if (value != NUMBER) {
      throw new IllegalArgumentException("protocolVersion must equal " + NUMBER);
    }
  }

  @Override
  public Integer scalarValue() {
    return value;
  }

  @Override
  public String toString() {
    return Integer.toString(value);
  }
}
