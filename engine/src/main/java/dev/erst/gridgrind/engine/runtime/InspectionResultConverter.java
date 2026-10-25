package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.InspectionResult;

/** Thin facade over the split inspection-result conversion supports. */
final class InspectionResultConverter {
  private InspectionResultConverter() {}

  static InspectionResult toReadResult(dev.erst.gridgrind.excel.WorkbookReadResult result) {
    return InspectionResultReadConverter.toReadResult(result);
  }
}
