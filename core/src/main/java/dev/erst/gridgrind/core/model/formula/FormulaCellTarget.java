package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.vocabulary.ProtocolCellAddressValidation;
import dev.erst.gridgrind.core.vocabulary.SheetName;

/** One exact workbook-qualified formula cell used by targeted calculation policy. */
public record FormulaCellTarget(SheetName sheetName, String address) {
  public FormulaCellTarget {
    java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
    address = ProtocolCellAddressValidation.validateAddress(address);
  }
}
