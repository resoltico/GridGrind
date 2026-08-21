package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing border patch used by {@link CellStylePatchInput}. */
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
