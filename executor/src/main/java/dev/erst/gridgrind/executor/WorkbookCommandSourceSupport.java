package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Base64;

/** Resolves source-backed contract payloads into inline values before engine translation. */
final class WorkbookCommandSourceSupport {
  private WorkbookCommandSourceSupport() {}

  static String inlineText(TextSourceInput source, String fieldName) {
    if (source instanceof TextSourceInput.Inline inline) {
      return inline.text();
    }
    throw new IllegalStateException(fieldName + " must be resolved to INLINE before conversion");
  }

  static byte[] inlineBinary(BinarySourceInput source, String fieldName) {
    if (source instanceof BinarySourceInput.InlineBase64 inline) {
      return Base64.getDecoder().decode(inline.base64Data());
    }
    throw new IllegalStateException(
        fieldName + " must be resolved to INLINE_BASE64 before conversion");
  }
}
