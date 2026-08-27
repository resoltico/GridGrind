package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;

/** One exact workbook-qualified formula cell used by targeted calculation policy. */
public record FormulaCellTarget(String sheetName, String address) {
  public FormulaCellTarget {
    ExcelSheetNames.requireValid(sheetName, "sheetName");
    address = ProtocolCellAddressValidation.validateAddress(address);
  }
}
