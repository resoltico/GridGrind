package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.RequestWarning;
import java.util.List;
import java.util.Objects;

/** Decoded request text together with a UTF-16-index to UTF-8-byte-offset map. */
final class RequestUtf8DecodeResult {
  private final String text;
  private final long[] byteOffsets;
  private final List<RequestStructuralProblem> problems;
  private final List<RequestWarning> warnings;

  RequestUtf8DecodeResult(
      String text,
      long[] byteOffsets,
      List<RequestStructuralProblem> problems,
      List<RequestWarning> warnings) {
    this.text = Objects.requireNonNull(text, "text must not be null");
    this.byteOffsets = Objects.requireNonNull(byteOffsets, "byteOffsets must not be null").clone();
    this.problems = List.copyOf(Objects.requireNonNull(problems, "problems must not be null"));
    this.warnings = List.copyOf(Objects.requireNonNull(warnings, "warnings must not be null"));
  }

  String text() {
    return text;
  }

  List<RequestStructuralProblem> problems() {
    return problems;
  }

  List<RequestWarning> warnings() {
    return warnings;
  }

  long byteOffsetAt(int characterOffset) {
    int boundedOffset = Math.min(Math.max(characterOffset, 0), byteOffsets.length - 1);
    return byteOffsets[boundedOffset];
  }
}
