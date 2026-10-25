package dev.erst.gridgrind.contract.json;

/** One object property occurrence retained even when its key is duplicated. */
record RequestJsonMember(String name, long nameByteOffset, RequestJsonNode value) {
  RequestJsonMember {
    name = RequestJsonNodeSupport.requireValue(name, "name");
    nameByteOffset = RequestJsonNodeSupport.requireByteOffset(nameByteOffset);
    value = RequestJsonNodeSupport.requireValue(value, "value");
  }
}
