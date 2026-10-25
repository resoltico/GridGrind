package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;

/** One exact workbook-qualified formula cell used by targeted calculation policy. */
public record FormulaCellTarget(SheetName sheetName, CellAddress address) {
  public FormulaCellTarget {
    java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
    java.util.Objects.requireNonNull(address, "address must not be null");
  }
}
