package dev.erst.gridgrind.json.intake;

import java.util.List;

/** A tolerant JSON array. */
record RequestJsonArray(long byteOffset, List<RequestJsonNode> elements)
    implements RequestJsonNode {
  RequestJsonArray {
    byteOffset = RequestJsonNodeSupport.requireByteOffset(byteOffset);
    elements = List.copyOf(RequestJsonNodeSupport.requireValue(elements, "elements"));
  }
}
