package dev.erst.gridgrind.contract.json;

/** A JSON null literal. */
record RequestJsonNull(long byteOffset) implements RequestJsonNode {
  RequestJsonNull {
    byteOffset = RequestJsonNodeSupport.requireByteOffset(byteOffset);
  }
}
