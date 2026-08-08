package dev.erst.gridgrind.contract.json;

/** A JSON number literal retained in its authored decimal spelling. */
record RequestJsonNumber(long byteOffset, String value) implements RequestJsonNode {
  RequestJsonNumber {
    byteOffset = RequestJsonNodeSupport.requireByteOffset(byteOffset);
    value = RequestJsonNodeSupport.requireValue(value, "value");
  }
}
