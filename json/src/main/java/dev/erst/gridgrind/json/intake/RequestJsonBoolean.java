package dev.erst.gridgrind.json.intake;

/** A JSON boolean literal. */
record RequestJsonBoolean(long byteOffset, boolean value) implements RequestJsonNode {
  RequestJsonBoolean {
    byteOffset = RequestJsonNodeSupport.requireByteOffset(byteOffset);
  }
}
