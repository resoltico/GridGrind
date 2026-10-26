package dev.erst.gridgrind.core.source;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.Objects;

/** Contract-owned authored text source used by source-backed mutation payloads. */
@Union(discriminator = "type", catalog = true, displayOrder = 4)
public sealed interface TextSourceInput
    permits dev.erst.gridgrind.core.source.TextSourceInput.Inline,
        dev.erst.gridgrind.core.source.TextSourceInput.Utf8File,
        dev.erst.gridgrind.core.source.TextSourceInput.StandardInput {
  /** Returns an explicitly empty inline text source. */
  static TextSourceInput empty() {
    return new Inline("");
  }

  /** Returns one inline UTF-8 text source. */
  static Inline inline(String text) {
    return new Inline(text);
  }

  /** Returns one UTF-8 text-file source. */
  static Utf8File utf8File(String path) {
    return new Utf8File(path);
  }

  /** Returns one standard-input text source. */
  static StandardInput standardInput() {
    return new StandardInput();
  }

  /** Inline UTF-8 text embedded directly in the JSON contract. */
  @Variant(id = "INLINE", summary = "Embed UTF-8 text directly in the request JSON.")
  record Inline(String text) implements TextSourceInput {
    public Inline {
      Objects.requireNonNull(text, "text must not be null");
    }
  }

  /** UTF-8 text loaded from one file path in the execution environment. */
  @Variant(
      id = "UTF8_FILE",
      summary = "Load UTF-8 text from one file path.",
      noteRefs = {"requestOwnedPathRule"})
  record Utf8File(String path) implements TextSourceInput {
    public Utf8File {
      path = requireNonBlank(path, "path");
    }
  }

  /** Text loaded from the execution transport's standard input byte stream as UTF-8. */
  @Variant(
      id = "STANDARD_INPUT",
      summary = "Load UTF-8 text from the execution transport's bound standard input bytes.")
  record StandardInput() implements TextSourceInput {}

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
