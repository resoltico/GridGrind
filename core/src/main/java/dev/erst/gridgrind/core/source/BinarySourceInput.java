package dev.erst.gridgrind.core.source;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.Base64;
import java.util.Objects;

/** Contract-owned authored binary source used by source-backed payloads. */
@Union(discriminator = "type", catalog = true, displayOrder = 5)
public sealed interface BinarySourceInput
    permits dev.erst.gridgrind.core.source.BinarySourceInput.InlineBase64,
        dev.erst.gridgrind.core.source.BinarySourceInput.File,
        dev.erst.gridgrind.core.source.BinarySourceInput.StandardInput {

  /** Returns one inline base64-authored binary source. */
  static InlineBase64 inlineBase64(String base64Data) {
    return new InlineBase64(base64Data);
  }

  /** Returns one file-backed binary source. */
  static File file(String path) {
    return new File(path);
  }

  /** Returns one standard-input binary source. */
  static StandardInput standardInput() {
    return new StandardInput();
  }

  /** Inline base64-encoded binary content embedded directly in the JSON contract. */
  @Variant(
      id = "INLINE_BASE64",
      summary = "Embed base64-encoded binary content directly in the request JSON.")
  record InlineBase64(String base64Data) implements BinarySourceInput {
    public InlineBase64 {
      base64Data = requireBase64(base64Data, "base64Data");
    }
  }

  /** Binary payload loaded from one file path in the execution environment. */
  @Variant(
      id = "FILE",
      summary = "Load binary content from one file path.",
      noteRefs = {"requestOwnedPathRule"})
  record File(String path) implements BinarySourceInput {
    public File {
      path = requireNonBlank(path, "path");
    }
  }

  /** Binary payload loaded from the execution transport's standard input byte stream. */
  @Variant(
      id = "STANDARD_INPUT",
      summary = "Load binary content from the execution transport's bound standard input bytes.")
  record StandardInput() implements BinarySourceInput {}

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static String requireBase64(String value, String fieldName) {
    String validatedValue = requireNonBlank(value, fieldName);
    try {
      Base64.getDecoder().decode(validatedValue);
    } catch (IllegalArgumentException exception) {
      throw new IllegalArgumentException(fieldName + " must be valid base64", exception);
    }
    return validatedValue;
  }
}
