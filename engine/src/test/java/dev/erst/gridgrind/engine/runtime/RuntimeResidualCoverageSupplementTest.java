package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.dto.CellGridInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellRowInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCellValue;
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
  void executionJournalTargetLabelsCoverResidualCellKindsAndSourceShapes() {
    TableSelector.ByName table = new TableSelector.ByName("BudgetTable");

    ExecutionJournal.Target fileFormulaTarget =
        ExecutionJournalTargetResolver.summaryTarget(
            new TableRowSelector.ByKeyCell(
                table,
                "Owner",
                new CellInput.Formula(TextSourceInput.utf8File("formulas/amount.txt"))));
    ExecutionJournal.Target standardInputFormulaTarget =
        ExecutionJournalTargetResolver.summaryTarget(
            new TableRowSelector.ByKeyCell(
                table, "Owner", new CellInput.Formula(TextSourceInput.standardInput())));

    assertEquals(
        "RichText[runs=1]",
        ExecutionJournalTargetResolver.summarizeCellInput(
            new CellInput.RichText(
                List.of(new RichTextRunInput(TextSourceInput.inline("Ada"), Optional.empty())))));
    assertEquals(
        "Error[error=#REF!]",
        ExecutionJournalTargetResolver.summarizeCellInput(new CellInput.ErrorValue("#REF!")));
    assertEquals(
        "Date[value=2026-06-12]",
        ExecutionJournalTargetResolver.summarizeCellInput(
            new CellInput.Date(LocalDate.of(2026, 6, 12))));
    assertEquals(
        "DateTime[value=2026-06-12T09:30]",
        ExecutionJournalTargetResolver.summarizeCellInput(
            new CellInput.DateTime(LocalDateTime.of(2026, 6, 12, 9, 30))));
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
            new CellInput.Formula(TextSourceInput.utf8File("formulas/value.txt")));
    CellMutationAction.SetCell standardInputFormula =
        new CellMutationAction.SetCell(new CellInput.Formula(TextSourceInput.standardInput()));

    assertEquals(Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(fileBackedFormula));
    assertEquals(
        Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(standardInputFormula));
  }

  @Test
  void sourceBackedMutationResolutionLeavesCompactUntypedPayloadsInPlace() throws Exception {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            Path.of("tmp", "runtime-residual-supplement"),
            Path.of("tmp", "runtime-residual-supplement", ".gridgrind", "tmp"),
            "unused".getBytes(StandardCharsets.UTF_8));
    CellMutationAction.SetRange compactRange =
        new CellMutationAction.SetRange(new CellGridInput.TextRows(List.of(List.of("Ada"))));
    CellMutationAction.AppendRow compactRow =
        new CellMutationAction.AppendRow(new CellRowInput.TextValues(List.of("Ada")));

    assertSame(compactRange, SourceBackedMutationActionResolver.resolve(compactRange, bindings));
    assertSame(compactRow, SourceBackedMutationActionResolver.resolve(compactRow, bindings));
  }

  @Test
  void sourceBackedPlanResolutionPreservesErrorCellsAndConvertersMapThemToExcelErrors()
      throws Exception {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            Path.of("tmp", "runtime-residual-supplement"),
            Path.of("tmp", "runtime-residual-supplement", ".gridgrind", "tmp"));
    CellInput.ErrorValue errorValue = new CellInput.ErrorValue("#REF!");

    assertSame(errorValue, SourceBackedPlanResolver.resolveCellInput(errorValue, bindings));
    assertEquals(
        "#REF!",
        assertInstanceOf(
                ExcelCellValue.ErrorValue.class,
                WorkbookCommandCellInputConverter.toExcelCellValue(errorValue))
            .value());
  }

  @Test
  void semanticSelectorResolverTreatsUnsupportedKeyCellKindsAsNonMatches() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", ExcelCellValue.text("Ada"));
      ExcelCellSnapshot snapshot = workbook.sheet("Budget").cells().snapshotCell("A1");

      boolean matched =
          SemanticSelectorResolver.matchesKeyCell(snapshot, new CellInput.ErrorValue("#REF!"));

      assertFalse(matched);
    }
  }
}
