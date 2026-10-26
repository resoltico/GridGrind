package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.model.value.CellGridSpec;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellRowSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.result.ExecutionJournal;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
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
  void sourceBackedPlanResolutionPreservesErrorCellsAndConvertersMapThemToExcelErrors()
      throws Exception {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            Path.of("tmp", "runtime-residual-supplement"),
            Path.of("tmp", "runtime-residual-supplement", "temp-root"),
            ExecutionGrantTestSupport.noPublication());
    CellValueSpec.ErrorValue errorValue = new CellValueSpec.ErrorValue("#REF!");

    assertSame(errorValue, SourceBackedPlanResolver.resolveCellInput(errorValue, bindings));
    assertEquals(
        "#REF!",
        assertInstanceOf(
                CellValue.ErrorValue.class,
                WorkbookCommandCellValueSpecConverter.toResolvedCellWrite(errorValue))
            .value());
  }

  @Test
  void semanticSelectorResolverTreatsUnsupportedKeyCellKindsAsNonMatches() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new CellValue.Text("Ada"));
      CellReport snapshot = workbook.sheet("Budget").cells().snapshotCell("A1");

      boolean matched =
          SemanticSelectorResolver.matchesKeyCell(
              snapshot, new CellValueSpec.ErrorValue("#REF!"), false);

      assertFalse(matched);
    }
  }

  @Test
  void conditionalFormattingMutationRequiresExplicitRangeTargets() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                DefaultGridGrindRequestExecutorTestSupport.command(
                    new RangeSelector.AllOnSheet(new SheetName("Budget")),
                    new StructuredMutationAction.SetConditionalFormatting(
                        new ConditionalFormattingSpec(
                            List.of(
                                new ConditionalFormattingRuleSpec.FormulaRule(
                                    "A1>0", true, Optional.empty()))))));

    assertEquals("Mutation target does not match SetConditionalFormatting", failure.getMessage());
  }
}
