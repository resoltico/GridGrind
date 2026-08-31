package dev.erst.gridgrind.engine.runtime.parity;

import dev.erst.gridgrind.contract.dto.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.contract.query.SheetInspectionResult;

/** Adjusts factual source formatting into the current authoring contract's expected form. */
final class ConditionalFormattingParityExpectations {
  private ConditionalFormattingParityExpectations() {}

  static SheetInspectionResult.ConditionalFormattingResult withTop10StopBarrier(
      SheetInspectionResult.ConditionalFormattingResult source) {
    return new SheetInspectionResult.ConditionalFormattingResult(
        source.stepId(),
        source.sheetName(),
        source.conditionalFormattingBlocks().stream()
            .map(ConditionalFormattingParityExpectations::withTop10StopBarrier)
            .toList());
  }

  private static ConditionalFormattingEntryReport withTop10StopBarrier(
      ConditionalFormattingEntryReport source) {
    return new ConditionalFormattingEntryReport(
        source.ranges(),
        source.rules().stream()
            .map(ConditionalFormattingParityExpectations::withTop10StopBarrier)
            .toList());
  }

  private static ConditionalFormattingRuleReport withTop10StopBarrier(
      ConditionalFormattingRuleReport rule) {
    if (rule instanceof ConditionalFormattingRuleReport.Top10Rule top10) {
      return new ConditionalFormattingRuleReport.Top10Rule(
          top10.priority(), true, top10.rank(), top10.percent(), top10.bottom(), top10.style());
    }
    return rule;
  }
}
