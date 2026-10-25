package dev.erst.gridgrind.json.intake;

/** A recoverable placeholder for malformed JSON syntax. */
record RequestJsonInvalid(long byteOffset) implements RequestJsonNode {
  RequestJsonInvalid {
    byteOffset = RequestJsonNodeSupport.requireByteOffset(byteOffset);
  }
}
