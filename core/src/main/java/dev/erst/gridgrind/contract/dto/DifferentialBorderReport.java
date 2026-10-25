package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Stream;

/** Protocol-facing factual report for one conditional-formatting differential border patch. */
public record DifferentialBorderReport(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<DifferentialBorderSideReport> all,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<DifferentialBorderSideReport> top,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<DifferentialBorderSideReport> right,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<DifferentialBorderSideReport> bottom,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<DifferentialBorderSideReport> left) {
  public DifferentialBorderReport {
    Objects.requireNonNull(all, "all must not be null");
    Objects.requireNonNull(top, "top must not be null");
    Objects.requireNonNull(right, "right must not be null");
    Objects.requireNonNull(bottom, "bottom must not be null");
    Objects.requireNonNull(left, "left must not be null");
    if (hasNoSides(all, top, right, bottom, left)) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }

  private static boolean hasNoSides(Optional<?>... sides) {
    return Stream.of(sides).allMatch(Optional::isEmpty);
  }
}
