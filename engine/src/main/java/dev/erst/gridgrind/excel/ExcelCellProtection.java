package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;

/** Protection patch applied through {@link ExcelCellStyle}. */
public record ExcelCellProtection(Optional<Boolean> locked, Optional<Boolean> hiddenFormula) {
  public ExcelCellProtection {
    Objects.requireNonNull(locked, "locked must not be null");
    Objects.requireNonNull(hiddenFormula, "hiddenFormula must not be null");
    if (locked.isEmpty() && hiddenFormula.isEmpty()) {
      throw new IllegalArgumentException("protection must set at least one attribute");
    }
  }
}
