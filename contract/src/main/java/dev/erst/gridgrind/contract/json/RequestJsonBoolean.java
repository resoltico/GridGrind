package dev.erst.gridgrind.contract.json;

/** A JSON boolean literal. */
record RequestJsonBoolean(long byteOffset, boolean value) implements RequestJsonNode {
  RequestJsonBoolean {
    byteOffset = RequestJsonNodeSupport.requireByteOffset(byteOffset);
  }
}
