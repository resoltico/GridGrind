package dev.erst.gridgrind.json.intake;

import java.util.regex.Pattern;
import tools.jackson.databind.node.JsonNodeFactory;

/** Renders request fields without confusing authored keys with path separators or the root. */
final class RequestJsonPaths {
  private static final Pattern SIMPLE_FIELD = Pattern.compile("[A-Za-z_][A-Za-z0-9_]*");

  private RequestJsonPaths() {}

  static String field(String parent, String name) {
    if (SIMPLE_FIELD.matcher(name).matches() && !(parent.isEmpty() && "request".equals(name))) {
      return parent.isEmpty() ? name : parent + "." + name;
    }
    return parent + "[" + JsonNodeFactory.instance.stringNode(name) + "]";
  }
}
