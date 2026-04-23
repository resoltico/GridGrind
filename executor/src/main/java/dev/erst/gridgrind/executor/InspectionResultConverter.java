package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.query.InspectionResult;

/** Thin facade over the split inspection-result conversion supports. */
final class InspectionResultConverter {
  private InspectionResultConverter() {}

  static InspectionResult toReadResult(dev.erst.gridgrind.excel.WorkbookReadResult result) {
    return InspectionResultReadConverter.toReadResult(result);
  }
}
