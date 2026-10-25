package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.vocabulary.ExcelSheetNames;
import dev.erst.gridgrind.core.vocabulary.ProtocolCellAddressValidation;

/** One exact workbook-qualified formula cell used by targeted calculation policy. */
public record FormulaCellTarget(String sheetName, String address) {
  public FormulaCellTarget {
    ExcelSheetNames.requireValid(sheetName, "sheetName");
    address = ProtocolCellAddressValidation.validateAddress(address);
  }
}
