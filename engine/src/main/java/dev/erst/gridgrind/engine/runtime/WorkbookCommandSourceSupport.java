package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.util.Base64;

/** Resolves source-backed contract payloads into inline values before engine translation. */
final class WorkbookCommandSourceSupport {
  private WorkbookCommandSourceSupport() {}

  static String inlineText(TextSourceSpec source, String fieldName) {
    if (source instanceof TextSourceSpec.Inline inline) {
      return inline.text();
    }
    throw new IllegalStateException(fieldName + " must be resolved to INLINE before conversion");
  }

  static byte[] inlineBinary(BinarySourceSpec source, String fieldName) {
    if (source instanceof BinarySourceSpec.InlineBase64 inline) {
      return Base64.getDecoder().decode(inline.base64Data());
    }
    throw new IllegalStateException(
        fieldName + " must be resolved to INLINE_BASE64 before conversion");
  }
}
