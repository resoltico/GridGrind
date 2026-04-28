package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.formulaCell;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct branch coverage for executor-side diagnostic field extraction. */
class ExecutionDiagnosticFieldsTest {
  @Test
  void exceptionLocationsCoverNamedRangeFormulaCellSheetAddressRangeAndUnknownVariants() {
    assertEquals(
        new ProblemContext.ProblemLocation.SheetNamedRange("Budget", "BudgetTotal"),
        ExecutionDiagnosticFields.locationFor(
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.SheetScope("Budget"))));
    assertEquals(
        new ProblemContext.ProblemLocation.NamedRange("BudgetTotal"),
        ExecutionDiagnosticFields.locationFor(
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertEquals(
        new ProblemContext.ProblemLocation.FormulaCell("Budget", "B4", "SUM("),
        ExecutionDiagnosticFields.locationFor(
            new InvalidFormulaException("Budget", "B4", "SUM(", "bad formula", null)));
    assertEquals(
        new ProblemContext.ProblemLocation.Cell("Budget", "B4"),
        ExecutionDiagnosticFields.locationFor(
            new InvalidFormulaException("Budget", "B4", null, "bad formula", null)));
    assertEquals(
        new ProblemContext.ProblemLocation.Sheet("Budget"),
        ExecutionDiagnosticFields.locationFor(new SheetNotFoundException("Budget")));
    assertEquals(
        new ProblemContext.ProblemLocation.Address("B4"),
        ExecutionDiagnosticFields.locationFor(new InvalidCellAddressException("B4", null)));
    assertEquals(
        new ProblemContext.ProblemLocation.RangeOnly("A1:B2"),
        ExecutionDiagnosticFields.locationFor(new InvalidRangeAddressException("A1:B2", null)));
    assertEquals(
        ProblemContext.ProblemLocation.unknown(),
        ExecutionDiagnosticFields.locationFor(new RuntimeException("boom")));
  }

  @Test
  void mutationActionLocationsReflectReachableDiagnosticShapes() {
    MutationAction.SetNamedRange sheetScopedNamedRange =
        new MutationAction.SetNamedRange(
            "BudgetTotal",
            new NamedRangeScope.Sheet("Budget"),
            new NamedRangeTarget("Budget", "B4"));
    MutationAction.SetNamedRange workbookFormulaNamedRange =
        new MutationAction.SetNamedRange(
            "BudgetExpr",
            new NamedRangeScope.Workbook(),
            new NamedRangeTarget("SUM(Budget!B2:B4)"));
    MutationAction.SetPivotTable pivot =
        new MutationAction.SetPivotTable(
            new PivotTableInput(
                "SalesPivot",
                "Budget",
                new PivotTableInput.Source.Range("Budget", "A1:B5"),
                new PivotTableInput.Anchor("C5"),
                List.of(),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableInput.DataField(
                        "Amount",
                        dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction.SUM,
                        null,
                        null))));
    MutationAction.SetTable table =
        new MutationAction.SetTable(
            new TableInput("BudgetTable", "Budget", "A1:B4", false, new TableStyleInput.None()));
    MutationAction.SetConditionalFormatting conditionalFormatting =
        new MutationAction.SetConditionalFormatting(
            new ConditionalFormattingBlockInput(
                List.of("B2:B5"),
                List.of(new ConditionalFormattingRuleInput.FormulaRule("B2>0", false, null))));

    assertEquals(
        new ProblemContext.ProblemLocation.SheetNamedRange("Budget", "BudgetTotal"),
        ExecutionDiagnosticFields.locationFor(sheetScopedNamedRange));
    assertEquals(
        new ProblemContext.ProblemLocation.NamedRange("BudgetExpr"),
        ExecutionDiagnosticFields.locationFor(workbookFormulaNamedRange));
    assertEquals(
        new ProblemContext.ProblemLocation.Cell("Budget", "C5"),
        ExecutionDiagnosticFields.locationFor(pivot));
    assertEquals(
        new ProblemContext.ProblemLocation.Range("Budget", "A1:B4"),
        ExecutionDiagnosticFields.locationFor(table));
    assertEquals(
        new ProblemContext.ProblemLocation.RangeOnly("B2:B5"),
        ExecutionDiagnosticFields.locationFor(conditionalFormatting));
    assertEquals(
        ProblemContext.ProblemLocation.unknown(),
        ExecutionDiagnosticFields.locationFor(
            new MutationAction.SetCell(formulaCell("SUM(A1:A2)"))));
  }

  @Test
  void selectorLocationsAndWorkbookStepFallbacksStayTyped() {
    assertEquals(
        new ProblemContext.ProblemLocation.SheetNamedRange("Budget", "BudgetTotal"),
        ExecutionDiagnosticFields.locationFor(
            new NamedRangeSelector.SheetScope("BudgetTotal", "Budget")));
    assertEquals(
        new ProblemContext.ProblemLocation.NamedRange("BudgetTotal"),
        ExecutionDiagnosticFields.locationFor(new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionDiagnosticFields.namedRangeNameFor(
            new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        new ProblemContext.ProblemLocation.Cell("Budget", "B4"),
        ExecutionDiagnosticFields.locationFor(new CellSelector.ByAddress("Budget", "B4")));
    assertEquals(
        new ProblemContext.ProblemLocation.Range("Budget", "A1:B2"),
        ExecutionDiagnosticFields.locationFor(new RangeSelector.ByRange("Budget", "A1:B2")));
    assertEquals(
        new ProblemContext.ProblemLocation.Sheet("Budget"),
        ExecutionDiagnosticFields.locationFor(new SheetSelector.ByName("Budget")));
    assertEquals(
        ProblemContext.ProblemLocation.unknown(),
        ExecutionDiagnosticFields.locationFor(new WorkbookSelector.Current()));
    assertEquals(
        ProblemContext.ProblemLocation.unknown(),
        ExecutionDiagnosticFields.locationFor(new Assertion.TablePresent()));

    MutationStep pivotAddressStep =
        new MutationStep(
            "set-pivot",
            new PivotTableSelector.ByNameOnSheet("SalesPivot", "Budget"),
            new MutationAction.SetPivotTable(
                new PivotTableInput(
                    "SalesPivot",
                    "Budget",
                    new PivotTableInput.Source.Range("Budget", "A1:B5"),
                    new PivotTableInput.Anchor("C5"),
                    List.of(),
                    List.of(),
                    List.of(),
                    List.of(
                        new PivotTableInput.DataField(
                            "Amount",
                            dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction
                                .SUM,
                            null,
                            null)))));
    MutationStep pivotNamedRangeStep =
        new MutationStep(
            "set-pivot-named-range",
            new PivotTableSelector.ByNameOnSheet("SalesPivot", "Budget"),
            new MutationAction.SetPivotTable(
                new PivotTableInput(
                    "SalesPivot",
                    "Budget",
                    new PivotTableInput.Source.NamedRange("BudgetSource"),
                    new PivotTableInput.Anchor("C5"),
                    List.of(),
                    List.of(),
                    List.of(),
                    List.of(
                        new PivotTableInput.DataField(
                            "Amount",
                            dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction
                                .SUM,
                            null,
                            null)))));
    MutationStep tableStep =
        new MutationStep(
            "set-table",
            new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
            new MutationAction.SetTable(
                new TableInput(
                    "BudgetTable", "Budget", "A1:B4", false, new TableStyleInput.None())));
    MutationStep formulaNamedRangeStep =
        new MutationStep(
            "set-name",
            new NamedRangeSelector.WorkbookScope("BudgetExpr"),
            new MutationAction.SetNamedRange(
                "BudgetExpr",
                new NamedRangeScope.Workbook(),
                new NamedRangeTarget("SUM(Budget!B2:B4)")));
    AssertionStep formulaAssertionStep =
        new AssertionStep(
            "assert-formula",
            new CellSelector.ByAddress("Budget", "B4"),
            new Assertion.FormulaText("SUM(A1:A2)"));
    InspectionStep workbookInspection =
        new InspectionStep(
            "summary", new WorkbookSelector.Current(), new InspectionQuery.GetWorkbookSummary());

    assertEquals(Optional.of("Budget"), ExecutionDiagnosticFields.sheetNameFor(pivotAddressStep));
    assertEquals(Optional.of("C5"), ExecutionDiagnosticFields.addressFor(pivotAddressStep));
    assertEquals(
        Optional.of("BudgetExpr"),
        ExecutionDiagnosticFields.namedRangeNameFor(formulaNamedRangeStep));
    assertEquals(
        Optional.of("BudgetSource"),
        ExecutionDiagnosticFields.namedRangeNameFor(pivotNamedRangeStep));
    assertEquals(Optional.of("A1:B4"), ExecutionDiagnosticFields.rangeFor(tableStep));
    assertEquals(
        Optional.of("SUM(A1:A2)"), ExecutionDiagnosticFields.formulaFor(formulaAssertionStep));
    assertEquals(
        Optional.of("Budget"), ExecutionDiagnosticFields.sheetNameFor(formulaAssertionStep));
    assertEquals(Optional.of("B4"), ExecutionDiagnosticFields.addressFor(formulaAssertionStep));
    assertEquals(Optional.empty(), ExecutionDiagnosticFields.rangeFor(formulaAssertionStep));
    assertEquals(
        Optional.empty(), ExecutionDiagnosticFields.namedRangeNameFor(formulaAssertionStep));
    assertEquals(Optional.empty(), ExecutionDiagnosticFields.addressFor(workbookInspection));
    assertEquals(Optional.empty(), ExecutionDiagnosticFields.rangeFor(workbookInspection));
    assertEquals(Optional.empty(), ExecutionDiagnosticFields.namedRangeNameFor(workbookInspection));
    assertEquals(
        Optional.of("Budget"),
        ExecutionDiagnosticFields.sheetNameFor(
            workbookInspection, new SheetNotFoundException("Budget")));
    assertEquals(
        Optional.of("B4"),
        ExecutionDiagnosticFields.addressFor(
            workbookInspection, new InvalidCellAddressException("B4", null)));
    assertEquals(
        Optional.of("A1:B2"),
        ExecutionDiagnosticFields.rangeFor(
            workbookInspection, new InvalidRangeAddressException("A1:B2", null)));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionDiagnosticFields.namedRangeNameFor(
            workbookInspection,
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertEquals(
        Optional.of("SUM("),
        ExecutionDiagnosticFields.formulaFor(
            workbookInspection, new InvalidFormulaException("Budget", "B4", "SUM(", "bad", null)));
    assertEquals(
        Optional.of("A1:B2"),
        ExecutionDiagnosticFields.rangeFor(new InvalidRangeAddressException("A1:B2", null)));
    assertEquals(
        Optional.of("SUM("),
        ExecutionDiagnosticFields.formulaFor(
            new InvalidFormulaException("Budget", "B4", "SUM(", "bad", null)));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionDiagnosticFields.namedRangeNameFor(
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertInstanceOf(
        ProblemContext.ProblemLocation.FormulaCell.class,
        ExecutionDiagnosticFields.locationFor(
            new MutationStep(
                "set-formula",
                new CellSelector.ByAddress("Budget", "B4"),
                new MutationAction.SetCell(formulaCell("SUM(A1:A2)"))),
            new InvalidFormulaException("Budget", "B4", "SUM(A1:A2)", "bad", null)));
  }
}
