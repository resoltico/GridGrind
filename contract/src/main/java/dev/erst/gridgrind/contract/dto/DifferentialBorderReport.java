package dev.erst.gridgrind.contract.dto;

import java.util.Objects;
import java.util.stream.Stream;

/** Protocol-facing factual report for one conditional-formatting differential border patch. */
public record DifferentialBorderReport(
    DifferentialBorderSideReport all,
    DifferentialBorderSideReport top,
    DifferentialBorderSideReport right,
    DifferentialBorderSideReport bottom,
    DifferentialBorderSideReport left) {
  public DifferentialBorderReport {
    if (hasNoSides(all, top, right, bottom, left)) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }

  private static boolean hasNoSides(Object... sides) {
    return Stream.of(sides).allMatch(Objects::isNull);
  }
}
