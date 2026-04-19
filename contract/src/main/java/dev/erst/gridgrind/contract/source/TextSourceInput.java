package dev.erst.gridgrind.contract.source;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Contract-owned authored text source used by source-backed mutation payloads. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = TextSourceInput.Inline.class, name = "INLINE"),
  @JsonSubTypes.Type(value = TextSourceInput.Utf8File.class, name = "UTF8_FILE"),
  @JsonSubTypes.Type(value = TextSourceInput.StandardInput.class, name = "STANDARD_INPUT")
})
public sealed interface TextSourceInput
    permits TextSourceInput.Inline, TextSourceInput.Utf8File, TextSourceInput.StandardInput {

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
  record Inline(String text) implements TextSourceInput {
    public Inline {
      Objects.requireNonNull(text, "text must not be null");
    }
  }

  /** UTF-8 text loaded from one file path in the execution environment. */
  record Utf8File(String path) implements TextSourceInput {
    public Utf8File {
      path = requireNonBlank(path, "path");
    }
  }

  /** Text loaded from the execution transport's standard input byte stream as UTF-8. */
  record StandardInput() implements TextSourceInput {}

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
