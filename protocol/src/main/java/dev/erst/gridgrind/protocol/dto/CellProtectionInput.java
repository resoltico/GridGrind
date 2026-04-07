package dev.erst.gridgrind.protocol.dto;

/** Protocol-facing cell protection patch used by {@link CellStyleInput}. */
public record CellProtectionInput(Boolean locked, Boolean hiddenFormula) {
  public CellProtectionInput {
    if (locked == null && hiddenFormula == null) {
      throw new IllegalArgumentException("protection must set at least one attribute");
    }
  }
}
