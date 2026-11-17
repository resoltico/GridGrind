package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.value.CellGridSpec;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellRowSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.result.ExecutionJournal;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct residual coverage for runtime helpers left outside end-to-end executor suites. */
class RuntimeResidualCoverageSupplementTest {
  @Test
  void summaryPhaseStrippingPreservesUntimedStatesAndRemovesTimingFromFinishedPhases() {
    ExecutionJournal.Phase notStarted = ExecutionJournal.Phase.notStarted();
    ExecutionJournal.Phase notRequested = ExecutionJournal.Phase.notRequested();
    ExecutionJournal.Phase timedSuccess =
        ExecutionJournal.Phase.succeeded("2026-07-03T10:00:00Z", "2026-07-03T10:00:01Z", 1);
    ExecutionJournal.Phase timedFailure =
        ExecutionJournal.Phase.failed("2026-07-03T10:00:00Z", "2026-07-03T10:00:01Z", 1);

    assertSame(notStarted, ExecutionJournalRecorder.phaseWithoutTiming(notStarted));
    assertSame(notRequested, ExecutionJournalRecorder.phaseWithoutTiming(notRequested));
    assertTrue(
        assertInstanceOf(
                ExecutionJournal.Phase.Succeeded.class,
                ExecutionJournalRecorder.phaseWithoutTiming(timedSuccess))
            .timing()
            .isEmpty());
    assertTrue(
        assertInstanceOf(
                ExecutionJournal.Phase.Failed.class,
                ExecutionJournalRecorder.phaseWithoutTiming(timedFailure))
            .timing()
            .isEmpty());
  }

  @Test
  void executionJournalTargetLabelsCoverResidualCellKindsAndSourceShapes() {
    TableSelector.ByName table = new TableSelector.ByName(new TableName("BudgetTable"));

    ExecutionJournal.Target fileFormulaTarget =
        ExecutionJournalTargetResolver.summaryTarget(
            new TableRowSelector.ByKeyCell(
                table,
                "Owner",
                new CellValueSpec.Formula(TextSourceSpec.utf8File("formulas/amount.txt"))));
    ExecutionJournal.Target standardInputFormulaTarget =
        ExecutionJournalTargetResolver.summaryTarget(
            new TableRowSelector.ByKeyCell(
                table, "Owner", new CellValueSpec.Formula(TextSourceSpec.standardInput())));

    assertEquals(
        "RichText[runs=1]",
        ExecutionJournalTargetResolver.summarizeCellInput(
            new CellValueSpec.RichText(
                List.of(new RichTextRunSpec(TextSourceSpec.inline("Ada"), Optional.empty())))));
    assertEquals(
        "Error[error=#REF!]",
        ExecutionJournalTargetResolver.summarizeCellInput(new CellValueSpec.ErrorValue("#REF!")));
    assertEquals(
        "Date[value=2026-06-12]",
        ExecutionJournalTargetResolver.summarizeCellInput(
            new CellValueSpec.Date(LocalDate.of(2026, 6, 12))));
    assertEquals(
        "DateTime[value=2026-06-12T09:30]",
        ExecutionJournalTargetResolver.summarizeCellInput(
            new CellValueSpec.DateTime(LocalDateTime.of(2026, 6, 12, 9, 30))));
    assertEquals(
        "Row where Owner=Formula[path=formulas/amount.txt] in Table BudgetTable",
        fileFormulaTarget.label());
    assertEquals(
        "Row where Owner=Formula[source=STANDARD_INPUT] in Table BudgetTable",
        standardInputFormulaTarget.label());
  }

  @Test
  void executionActionDiagnosticFieldsOmitNonInlineFormulaSources() {
    CellMutationAction.SetCell fileBackedFormula =
        new CellMutationAction.SetCell(
            new CellValueSpec.Formula(TextSourceSpec.utf8File("formulas/value.txt")));
    CellMutationAction.SetCell standardInputFormula =
        new CellMutationAction.SetCell(new CellValueSpec.Formula(TextSourceSpec.standardInput()));

    assertEquals(Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(fileBackedFormula));
    assertEquals(
        Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(standardInputFormula));
  }

  @Test
  void sourceBackedMutationResolutionLeavesCompactUntypedPayloadsInPlace() throws Exception {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            Path.of("tmp", "runtime-residual-supplement"),
            Path.of("tmp", "runtime-residual-supplement", "temp-root"),
            "unused".getBytes(StandardCharsets.UTF_8),
            ExecutionGrantTestSupport.noPublicationWithStandardInput(List.of()));
    CellMutationAction.SetRange compactRange =
        new CellMutationAction.SetRange(new CellGridSpec.TextRows(List.of(List.of("Ada"))));
    CellMutationAction.AppendRow compactRow =
        new CellMutationAction.AppendRow(new CellRowSpec.TextValues(List.of("Ada")));

    assertSame(compactRange, SourceBackedMutationActionResolver.resolve(compactRange, bindings));
    assertSame(compactRow, SourceBackedMutationActionResolver.resolve(compactRow, bindings));
  }

  @Test
  void sourceBackedPlanResolutionPreservesErrorCells() throws Exception {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            Path.of("tmp", "runtime-residual-supplement"),
            Path.of("tmp", "runtime-residual-supplement", "temp-root"),
            ExecutionGrantTestSupport.noPublication());
    CellValueSpec.ErrorValue errorValue = new CellValueSpec.ErrorValue("#REF!");

    assertSame(errorValue, SourceBackedPlanResolver.resolveCellInput(errorValue, bindings));
  }

  @Test
  void semanticSelectorResolverTreatsUnsupportedKeyCellKindsAsNonMatches() throws Exception {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Budget")
          .cells()
          .setCell(
              new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
              new dev.erst.gridgrind.core.model.value.ResolvedCellText("Ada"));
      CellReport snapshot =
          workbook
              .sheet("Budget")
              .cells()
              .snapshotCell(new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"));

      boolean matched =
          TableCellResolver.matchesKeyCell(snapshot, new CellValueSpec.ErrorValue("#REF!"));

      assertFalse(matched);
    }
  }
}
