package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.*;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests request warnings derived from same-request spaced sheet-name formula references. */
class GridGrindRequestWarningsTest {
  @Test
  void warnsWhenSetCellFormulaReferencesSameRequestSpacedSheetWithoutQuotes() {
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                mutate(new SheetSelector.ByName("Budget Review"), new MutationAction.EnsureSheet()),
                mutate(new SheetSelector.ByName("Summary"), new MutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress("Summary", "A1"),
                    new MutationAction.SetCell(formulaCell("SUM(Budget Review!A1)")))));

    List<RequestWarning> warnings = GridGrindRequestWarnings.collect(request);

    assertEquals(1, warnings.size());
    assertEquals(2, warnings.getFirst().stepIndex());
    assertEquals("step-03-set-cell", warnings.getFirst().stepId());
    assertEquals("SET_CELL", warnings.getFirst().stepType());
    assertTrue(warnings.getFirst().message().contains("Budget Review"));
    assertTrue(warnings.getFirst().message().contains("'Sheet Name'!A1"));
  }

  @Test
  void doesNotWarnWhenSpacedSheetReferencesAreQuotedOrOnlyAppearInsideStrings() {
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                mutate(new SheetSelector.ByName("Budget Review"), new MutationAction.EnsureSheet()),
                mutate(new SheetSelector.ByName("Summary"), new MutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress("Summary", "A1"),
                    new MutationAction.SetCell(formulaCell("SUM('Budget Review'!A1)"))),
                mutate(
                    new CellSelector.ByAddress("Summary", "A2"),
                    new MutationAction.SetCell(formulaCell("INDIRECT(\"Budget Review!A1\")")))));

    assertEquals(List.of(), GridGrindRequestWarnings.collect(request));
  }

  @Test
  void doesNotWarnWhenAnotherSheetNameOnlyContainsTheSpacedNameAsASuffix() {
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                mutate(new SheetSelector.ByName("Budget Review"), new MutationAction.EnsureSheet()),
                mutate(new SheetSelector.ByName("Summary"), new MutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress("Summary", "A1"),
                    new MutationAction.SetCell(formulaCell("SUM(Annual Budget Review!A1)")))));

    assertEquals(List.of(), GridGrindRequestWarnings.collect(request));
  }

  @Test
  void warnsWhenRenameAndCopyOperationsIntroduceSpacedSheetNames() {
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                mutate(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.RenameSheet("Budget Review")),
                mutate(
                    new SheetSelector.ByName("Budget Review"),
                    new MutationAction.CopySheet(
                        "Annual Budget Review", new SheetCopyPosition.AppendAtEnd())),
                mutate(new SheetSelector.ByName("Summary"), new MutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress("Summary", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.Formula(
                            text("SUM(Budget Review!A1,Annual Budget Review!$A$1)"))))));

    assertEquals(
        List.of(
            new RequestWarning(
                4,
                "step-05-set-cell",
                "SET_CELL",
                "Formula references same-request sheet names with spaces without single quotes: Budget Review, Annual Budget Review. Use 'Sheet Name'!A1 syntax.")),
        GridGrindRequestWarnings.collect(request));
  }

  @Test
  void doesNotWarnWhenUnquotedSheetTokensDoNotStartACellReference() {
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                mutate(new SheetSelector.ByName("Budget Review"), new MutationAction.EnsureSheet()),
                mutate(new SheetSelector.ByName("Summary"), new MutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress("Summary", "A1"),
                    new MutationAction.SetCell(formulaCell("Budget Review!"))),
                mutate(
                    new CellSelector.ByAddress("Summary", "A2"),
                    new MutationAction.SetCell(formulaCell("T(Budget Review!1)"))),
                mutate(
                    new CellSelector.ByAddress("Summary", "A3"),
                    new MutationAction.SetCell(formulaCell("T(\"Budget Review!A1\"\" suffix\")"))),
                mutate(
                    new CellSelector.ByAddress("Summary", "A4"),
                    new MutationAction.SetCell(formulaCell("\"Budget Review!A1\"")))));

    assertEquals(List.of(), GridGrindRequestWarnings.collect(request));
  }

  @Test
  void collectsWarningsFromSetRangeAndAppendRowFormulaCells() {
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                mutate(new SheetSelector.ByName("Budget Review"), new MutationAction.EnsureSheet()),
                mutate(new SheetSelector.ByName("Summary"), new MutationAction.EnsureSheet()),
                mutate(
                    new RangeSelector.ByRange("Summary", "A1:B1"),
                    new MutationAction.SetRange(
                        List.of(List.of(formulaCell("Budget Review!A1"), textCell("ok"))))),
                mutate(
                    new SheetSelector.ByName("Summary"),
                    new MutationAction.AppendRow(
                        List.of(textCell("Total"), formulaCell("SUM(Budget Review!A1)"))))));

    List<RequestWarning> warnings = GridGrindRequestWarnings.collect(request);

    assertEquals(
        List.of(
            new RequestWarning(
                2,
                "step-03-set-range",
                "SET_RANGE",
                "Formula references same-request sheet names with spaces without single quotes: Budget Review. Use 'Sheet Name'!A1 syntax."),
            new RequestWarning(
                3,
                "step-04-append-row",
                "APPEND_ROW",
                "Formula references same-request sheet names with spaces without single quotes: Budget Review. Use 'Sheet Name'!A1 syntax.")),
        warnings);
  }

  @Test
  void doesNotWarnWhenFormulaSourceIsNotInline() {
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                mutate(new SheetSelector.ByName("Budget Review"), new MutationAction.EnsureSheet()),
                mutate(new SheetSelector.ByName("Summary"), new MutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress("Summary", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.Formula(
                            dev.erst.gridgrind.contract.source.TextSourceInput.utf8File(
                                "formula.txt"))))));

    assertEquals(List.of(), GridGrindRequestWarnings.collect(request));
  }
}
