package dev.erst.gridgrind.contract.json;

import java.util.List;

/** A tolerant JSON object that preserves every authored property occurrence. */
record RequestJsonObject(long byteOffset, List<RequestJsonMember> members)
    implements RequestJsonNode {
  RequestJsonObject {
    byteOffset = RequestJsonNodeSupport.requireByteOffset(byteOffset);
    members = List.copyOf(RequestJsonNodeSupport.requireValue(members, "members"));
  }
}
