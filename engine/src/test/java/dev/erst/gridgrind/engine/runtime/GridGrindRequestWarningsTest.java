package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.value.CellInput;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
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
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget Review")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new SheetSelector.ByName(new SheetName("Summary")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress(new SheetName("Summary"), "A1"),
                    new CellMutationAction.SetCell(formulaCell("SUM(Budget Review!A1)")))));

    List<RequestWarning> warnings = GridGrindRequestWarnings.collect(request);

    assertEquals(1, warnings.size());
    dev.erst.gridgrind.core.problem.RequestWarningLocation.Step location =
        assertInstanceOf(
            dev.erst.gridgrind.core.problem.RequestWarningLocation.Step.class,
            warnings.getFirst().location());
    assertEquals(2, location.stepIndex());
    assertEquals("step-03-set-cell", location.stepId().value());
    assertEquals("SET_CELL", location.stepType());
    assertTrue(warnings.getFirst().message().contains("Budget Review"));
    assertTrue(warnings.getFirst().message().contains("'Sheet Name'!A1"));
  }

  @Test
  void ignoresAnIncompatibleButStructurallyBoundEnsureSheetTarget() {
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                mutate(new WorkbookSelector.Current(), new WorkbookMutationAction.EnsureSheet())));

    assertEquals(List.of(), GridGrindRequestWarnings.collect(request));
  }

  @Test
  void doesNotWarnWhenSpacedSheetReferencesAreQuotedOrOnlyAppearInsideStrings() {
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget Review")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new SheetSelector.ByName(new SheetName("Summary")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress(new SheetName("Summary"), "A1"),
                    new CellMutationAction.SetCell(formulaCell("SUM('Budget Review'!A1)"))),
                mutate(
                    new CellSelector.ByAddress(new SheetName("Summary"), "A2"),
                    new CellMutationAction.SetCell(
                        formulaCell("INDIRECT(\"Budget Review!A1\")")))));

    assertEquals(List.of(), GridGrindRequestWarnings.collect(request));
  }

  @Test
  void doesNotWarnWhenAnotherSheetNameOnlyContainsTheSpacedNameAsASuffix() {
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget Review")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new SheetSelector.ByName(new SheetName("Summary")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress(new SheetName("Summary"), "A1"),
                    new CellMutationAction.SetCell(formulaCell("SUM(Annual Budget Review!A1)")))));

    assertEquals(List.of(), GridGrindRequestWarnings.collect(request));
  }

  @Test
  void warnsWhenRenameAndCopyOperationsIntroduceSpacedSheetNames() {
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.RenameSheet(new SheetName("Budget Review"))),
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget Review")),
                    new WorkbookMutationAction.CopySheet(
                        new SheetName("Annual Budget Review"),
                        new SheetCopyPosition.AppendAtEnd())),
                mutate(
                    new SheetSelector.ByName(new SheetName("Summary")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress(new SheetName("Summary"), "A1"),
                    new CellMutationAction.SetCell(
                        new CellInput.Formula(
                            dev.erst.gridgrind.core.source.TextSourceInput.inline(
                                "SUM(Budget Review!A1,Annual Budget Review!$A$1)"))))));

    assertEquals(
        List.of(
            new RequestWarning(
                dev.erst.gridgrind.core.problem.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
                4,
                new StepId("step-05-set-cell"),
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
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget Review")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new SheetSelector.ByName(new SheetName("Summary")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress(new SheetName("Summary"), "A1"),
                    new CellMutationAction.SetCell(formulaCell("Budget Review!"))),
                mutate(
                    new CellSelector.ByAddress(new SheetName("Summary"), "A2"),
                    new CellMutationAction.SetCell(formulaCell("T(Budget Review!1)"))),
                mutate(
                    new CellSelector.ByAddress(new SheetName("Summary"), "A3"),
                    new CellMutationAction.SetCell(
                        formulaCell("T(\"Budget Review!A1\"\" suffix\")"))),
                mutate(
                    new CellSelector.ByAddress(new SheetName("Summary"), "A4"),
                    new CellMutationAction.SetCell(formulaCell("\"Budget Review!A1\"")))));

    assertEquals(List.of(), GridGrindRequestWarnings.collect(request));
  }

  @Test
  void collectsWarningsFromSetRangeAndAppendRowFormulaCells() {
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget Review")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new SheetSelector.ByName(new SheetName("Summary")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new RangeSelector.ByRange(new SheetName("Summary"), "A1:B1"),
                    new CellMutationAction.SetRange(
                        new dev.erst.gridgrind.core.model.value.CellGridInput.Typed(
                            List.of(List.of(formulaCell("Budget Review!A1"), textCell("ok")))))),
                mutate(
                    new SheetSelector.ByName(new SheetName("Summary")),
                    new CellMutationAction.AppendRow(
                        new dev.erst.gridgrind.core.model.value.CellRowInput.Typed(
                            List.of(textCell("Total"), formulaCell("SUM(Budget Review!A1)")))))));

    List<RequestWarning> warnings = GridGrindRequestWarnings.collect(request);

    assertEquals(
        List.of(
            new RequestWarning(
                dev.erst.gridgrind.core.problem.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
                2,
                new StepId("step-03-set-range"),
                "SET_RANGE",
                "Formula references same-request sheet names with spaces without single quotes: Budget Review. Use 'Sheet Name'!A1 syntax."),
            new RequestWarning(
                dev.erst.gridgrind.core.problem.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
                3,
                new StepId("step-04-append-row"),
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
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget Review")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new SheetSelector.ByName(new SheetName("Summary")),
                    new WorkbookMutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress(new SheetName("Summary"), "A1"),
                    new CellMutationAction.SetCell(
                        new CellInput.Formula(
                            dev.erst.gridgrind.core.source.TextSourceInput.utf8File(
                                "formula.txt"))))));

    assertEquals(List.of(), GridGrindRequestWarnings.collect(request));
  }
}
