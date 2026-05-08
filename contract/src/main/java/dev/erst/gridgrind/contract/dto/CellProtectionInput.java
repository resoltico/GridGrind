package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing cell protection patch used by {@link CellStyleInput}. */
public record CellProtectionInput(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Boolean> locked,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Boolean> hiddenFormula) {
  public CellProtectionInput {
    Objects.requireNonNull(locked, "locked must not be null");
    Objects.requireNonNull(hiddenFormula, "hiddenFormula must not be null");
    if (locked.isEmpty() && hiddenFormula.isEmpty()) {
      throw new IllegalArgumentException("protection must set at least one attribute");
    }
  }
}
