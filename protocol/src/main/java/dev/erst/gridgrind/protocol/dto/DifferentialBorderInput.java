package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;
import java.util.stream.Stream;

/** Protocol-facing differential border patch used by conditional-formatting rule styles. */
public record DifferentialBorderInput(
    DifferentialBorderSideInput all,
    DifferentialBorderSideInput top,
    DifferentialBorderSideInput right,
    DifferentialBorderSideInput bottom,
    DifferentialBorderSideInput left) {
  public DifferentialBorderInput {
    if (hasNoSides(all, top, right, bottom, left)) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }

  private static boolean hasNoSides(Object... sides) {
    return Stream.of(sides).allMatch(Objects::isNull);
  }
}
