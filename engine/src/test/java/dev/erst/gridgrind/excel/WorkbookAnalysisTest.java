package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookAnalysis payload invariants and defensive copies. */
class WorkbookAnalysisTest {
  @Test
  void enforcesLocationAndSummaryInvariants() {
    assertDoesNotThrow(() -> new WorkbookAnalysis.AnalysisLocation.Workbook());
    assertDoesNotThrow(() -> new WorkbookAnalysis.AnalysisLocation.Sheet("Budget"));
    assertDoesNotThrow(() -> new WorkbookAnalysis.AnalysisLocation.Cell("Budget", "A1"));
    assertDoesNotThrow(() -> new WorkbookAnalysis.AnalysisLocation.Range("Budget", "A1:B2"));
    assertDoesNotThrow(
        () ->
            new WorkbookAnalysis.AnalysisLocation.NamedRange(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope()));

    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookAnalysis.AnalysisLocation.Sheet(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookAnalysis.AnalysisLocation.Range("Budget", " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookAnalysis.AnalysisLocation.NamedRange(
                " ", new ExcelNamedRangeScope.WorkbookScope()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnalysis.AnalysisLocation.NamedRange("BudgetTotal", null));

    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookAnalysis.AnalysisSummary(-1, 0, 0, 0));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookAnalysis.AnalysisSummary(1, -1, 1, 1));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookAnalysis.AnalysisSummary(1, 1, -1, 1));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookAnalysis.AnalysisSummary(1, 1, 1, -1));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookAnalysis.AnalysisSummary(2, 1, 0, 0));
  }

  @Test
  void copiesFindingCollectionsAndRejectsInvalidAnalysisCounts() {
    WorkbookAnalysis.AnalysisFinding finding =
        new WorkbookAnalysis.AnalysisFinding(
            WorkbookAnalysis.AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
            WorkbookAnalysis.AnalysisSeverity.INFO,
            "Volatile formula function",
            "Formula uses a volatile function.",
            new WorkbookAnalysis.AnalysisLocation.Workbook(),
            List.of("NOW()"));
    WorkbookAnalysis.AnalysisFinding warningFinding =
        new WorkbookAnalysis.AnalysisFinding(
            WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET,
            WorkbookAnalysis.AnalysisSeverity.WARNING,
            "Malformed hyperlink target",
            "Hyperlink target is malformed.",
            new WorkbookAnalysis.AnalysisLocation.Sheet("Links"),
            List.of("mailto:"));
    WorkbookAnalysis.AnalysisFinding errorFinding =
        new WorkbookAnalysis.AnalysisFinding(
            WorkbookAnalysis.AnalysisFindingCode.NAMED_RANGE_BROKEN_REFERENCE,
            WorkbookAnalysis.AnalysisSeverity.ERROR,
            "Broken named range",
            "Named range contains #REF!.",
            new WorkbookAnalysis.AnalysisLocation.Range("Budget", "A1:B2"),
            List.of("#REF!"));
    WorkbookAnalysis.AnalysisSummary summary = new WorkbookAnalysis.AnalysisSummary(3, 1, 1, 1);

    List<WorkbookAnalysis.AnalysisFinding> mutable =
        new ArrayList<>(List.of(finding, warningFinding, errorFinding));
    WorkbookAnalysis.FormulaHealth formulaHealth =
        new WorkbookAnalysis.FormulaHealth(2, summary, mutable);
    WorkbookAnalysis.HyperlinkHealth hyperlinkHealth =
        new WorkbookAnalysis.HyperlinkHealth(2, summary, mutable);
    WorkbookAnalysis.NamedRangeHealth namedRangeHealth =
        new WorkbookAnalysis.NamedRangeHealth(2, summary, mutable);
    WorkbookAnalysis.WorkbookFindings workbookFindings =
        new WorkbookAnalysis.WorkbookFindings(summary, mutable);
    mutable.clear();

    assertEquals(3, formulaHealth.findings().size());
    assertEquals(3, hyperlinkHealth.findings().size());
    assertEquals(3, namedRangeHealth.findings().size());
    assertEquals(3, workbookFindings.findings().size());

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookAnalysis.FormulaHealth(-1, summary, List.of(finding)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookAnalysis.HyperlinkHealth(-1, summary, List.of(finding)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookAnalysis.NamedRangeHealth(-1, summary, List.of(finding)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookAnalysis.DataValidationHealth(-1, summary, List.of(finding)));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnalysis.FormulaHealth(1, null, List.of(finding)));
    assertThrows(
        NullPointerException.class, () -> new WorkbookAnalysis.HyperlinkHealth(1, summary, null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnalysis.NamedRangeHealth(1, summary, List.of(finding, null)));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnalysis.WorkbookFindings(null, List.of(finding)));
  }
}
