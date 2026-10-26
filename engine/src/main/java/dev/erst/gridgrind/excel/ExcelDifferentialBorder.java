package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Border patch used by conditional-formatting differential styles. */
public record ExcelDifferentialBorder(
    @Nullable BorderSidePatch all,
    @Nullable BorderSidePatch top,
    @Nullable BorderSidePatch right,
    @Nullable BorderSidePatch bottom,
    @Nullable BorderSidePatch left) {
  public ExcelDifferentialBorder {
    if (java.util.stream.Stream.of(all, top, right, bottom, left).allMatch(Objects::isNull)) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }
}
