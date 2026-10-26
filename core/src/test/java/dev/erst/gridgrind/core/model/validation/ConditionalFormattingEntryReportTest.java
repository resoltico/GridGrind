package dev.erst.gridgrind.core.model.validation;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.result.AnalysisSummaryReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing conditional-formatting factual report conversion. */
class ConditionalFormattingEntryReportTest {
  @Test
  void validatesConditionalFormattingHealthReport() {
    ConditionalFormattingHealthReport report =
        new ConditionalFormattingHealthReport(
            3,
            new AnalysisSummaryReport(1, 0, 1, 0),
            List.of(
                new AnalysisFindingReport(
                    AnalysisFindingCode.CONDITIONAL_FORMATTING_PRIORITY_COLLISION,
                    AnalysisSeverity.WARNING,
                    "Priority collision",
                    "Conditional-formatting priorities collide.",
                    new AnalysisLocationReport.Sheet(new SheetLabel("Ops")),
                    List.of("FORMULA_RULE@Ops!A1:A3"))));

    assertEquals(3, report.checkedConditionalFormattingBlockCount());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingEntryReport(
                List.of(new RangeLabel(" ")),
                List.of(
                    new ConditionalFormattingRuleReport.UnsupportedRule(
                        1, false, "FORMULA", "Missing formula text"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingHealthReport(
                -1, new AnalysisSummaryReport(0, 0, 0, 0), List.of()));
  }
}
