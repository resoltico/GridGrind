package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelSheetNames;
import java.util.Objects;

/** One concrete formula-cell target used by targeted evaluation operations. */
public record FormulaCellTargetInput(String sheetName, String address) {
  public FormulaCellTargetInput {
    ExcelSheetNames.requireValid(sheetName, "sheetName");
    Objects.requireNonNull(address, "address must not be null");
    if (address.isBlank()) {
      throw new IllegalArgumentException("address must not be blank");
    }
  }
}
