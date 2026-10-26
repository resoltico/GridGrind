package dev.erst.gridgrind.core.model.style;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing border patch used by {@link CellStylePatchInput}. */
@CatalogType(
    id = "CellBorderInput",
    summary =
        "Border patch for cell styling; at least one side must be set. Use 'all' as shorthand for all four sides.",
    displayOrder = 49)
public record CellBorderInput(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<BorderSideInput> all,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<BorderSideInput> top,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<BorderSideInput> right,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<BorderSideInput> bottom,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<BorderSideInput> left) {
  public CellBorderInput {
    Objects.requireNonNull(all, "all must not be null");
    Objects.requireNonNull(top, "top must not be null");
    Objects.requireNonNull(right, "right must not be null");
    Objects.requireNonNull(bottom, "bottom must not be null");
    Objects.requireNonNull(left, "left must not be null");
    if (all.isEmpty() && top.isEmpty() && right.isEmpty() && bottom.isEmpty() && left.isEmpty()) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }
}
