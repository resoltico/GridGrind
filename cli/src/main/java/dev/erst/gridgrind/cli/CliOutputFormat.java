package dev.erst.gridgrind.cli;

import java.util.Locale;
import java.util.Objects;

/** Global rendering mode for CLI-owned output surfaces. */
enum CliOutputFormat {
  TEXT,
  STRUCTURED;

  static CliOutputFormat parse(String rawValue) {
    String normalized =
        Objects.requireNonNull(rawValue, "rawValue must not be null")
            .trim()
            .toUpperCase(Locale.ROOT);
    return switch (normalized) {
      case "TEXT" -> TEXT;
      case "STRUCTURED" -> STRUCTURED;
      default ->
          throw new CliArgumentsException("--format", "--format must be one of: text, structured");
    };
  }
}
