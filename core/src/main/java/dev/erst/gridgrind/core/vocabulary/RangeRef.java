package dev.erst.gridgrind.core.vocabulary;

import java.util.Objects;

/** One local A1 cell or rectangular range, preserving authored endpoint order and spelling. */
public record RangeRef(String value) implements WireString {
  public RangeRef {
    Objects.requireNonNull(value, "range must not be null");
    String[] endpoints = value.split(":", -1);
    if (endpoints.length > 2) {
      throw new IllegalArgumentException("range must contain one cell or two endpoints");
    }
    for (String endpoint : endpoints) {
      ProtocolCellAddressValidation.validateAddress(endpoint);
    }
  }

  @Override
  public String toString() {
    return value;
  }
}
