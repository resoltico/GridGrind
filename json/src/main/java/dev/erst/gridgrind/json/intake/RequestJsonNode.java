package dev.erst.gridgrind.json.intake;

/** One syntax node retained by tolerant request parsing. */
sealed interface RequestJsonNode
    permits RequestJsonObject,
        RequestJsonArray,
        RequestJsonString,
        RequestJsonNumber,
        RequestJsonBoolean,
        RequestJsonNull,
        RequestJsonInvalid {

  /** Returns the zero-based UTF-8 byte offset at which this value begins. */
  long byteOffset();
}
