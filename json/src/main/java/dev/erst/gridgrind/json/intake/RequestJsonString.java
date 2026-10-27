package dev.erst.gridgrind.json.intake;

/** A decoded JSON string literal. */
record RequestJsonString(long byteOffset, String value) implements RequestJsonNode {
  RequestJsonString {
    byteOffset = RequestJsonNodeSupport.requireByteOffset(byteOffset);
    value = RequestJsonNodeSupport.requireValue(value, "value");
  }
}
