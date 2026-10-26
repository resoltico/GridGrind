package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.formulaCell;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.pivot.PivotTableInput;
import dev.erst.gridgrind.core.model.table.TableInput;
import dev.erst.gridgrind.core.model.table.TableStyleInput;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingDefinitionInput;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
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
        new ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange(
            new SheetLabel("Budget"), "BudgetTotal"),
        ExecutionDiagnosticFields.locationFor(
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.SheetScope("Budget"))));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.NamedRange("BudgetTotal"),
        ExecutionDiagnosticFields.locationFor(
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell(
            new SheetLabel("Budget"), "B4", "SUM("),
        ExecutionDiagnosticFields.locationFor(
            new InvalidFormulaException("Budget", "B4", "SUM(", "bad formula", null)));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Cell(new SheetLabel("Budget"), "B4"),
        ExecutionDiagnosticFields.locationFor(
            new InvalidFormulaException("Budget", "B4", null, "bad formula", null)));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Sheet(new SheetLabel("Budget")),
        ExecutionDiagnosticFields.locationFor(new SheetNotFoundException("Budget")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Address("B4"),
        ExecutionDiagnosticFields.locationFor(new InvalidCellAddressException("B4", null)));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.RangeOnly("A1:B2"),
        ExecutionDiagnosticFields.locationFor(new InvalidRangeAddressException("A1:B2", null)));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.unknown(),
        ExecutionDiagnosticFields.locationFor(new RuntimeException("boom")));
  }

  @Test
  void mutationActionLocationsReflectReachableDiagnosticShapes() {
    StructuredMutationAction.SetNamedRange sheetScopedNamedRange =
        new StructuredMutationAction.SetNamedRange(
            new DefinedName("BudgetTotal"),
            new NamedRangeScope.Sheet(new SheetName("Budget")),
            NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("B4")));
    StructuredMutationAction.SetNamedRange workbookFormulaNamedRange =
        new StructuredMutationAction.SetNamedRange(
            new DefinedName("BudgetExpr"),
            new NamedRangeScope.Workbook(),
            NamedRangeTarget.formula("SUM(Budget!B2:B4)"));
    StructuredMutationAction.SetPivotTable pivot =
        new StructuredMutationAction.SetPivotTable(
            new PivotTableInput(
                "SalesPivot",
                new SheetName("Budget"),
                new PivotTableInput.PivotSourceSpec.Range(
                    new SheetName("Budget"), new RangeRef("A1:B5")),
                new PivotTableInput.Anchor(new CellAddress("C5")),
                List.of(),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableInput.DataField(
                        "Amount",
                        dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction.SUM,
                        "Amount",
                        Optional.empty()))));
    StructuredMutationAction.SetTable table =
        new StructuredMutationAction.SetTable(
            TableInput.withDefaultMetadata(
                new TableName("BudgetTable"),
                new SheetName("Budget"),
                new RangeRef("A1:B4"),
                false,
                new TableStyleInput.None()));
    StructuredMutationAction.SetConditionalFormatting conditionalFormatting =
        new StructuredMutationAction.SetConditionalFormatting(
            new ConditionalFormattingDefinitionInput(
                List.of(
                    new ConditionalFormattingRuleInput.FormulaRule(
                        "B2>0", true, Optional.empty()))));

    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange(
            new SheetLabel("Budget"), "BudgetTotal"),
        ExecutionDiagnosticFields.locationFor(sheetScopedNamedRange));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.NamedRange("BudgetExpr"),
        ExecutionDiagnosticFields.locationFor(workbookFormulaNamedRange));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Cell(new SheetLabel("Budget"), "C5"),
        ExecutionDiagnosticFields.locationFor(pivot));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Range(new SheetLabel("Budget"), "A1:B4"),
        ExecutionDiagnosticFields.locationFor(table));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.unknown(),
        ExecutionDiagnosticFields.locationFor(conditionalFormatting));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.unknown(),
        ExecutionDiagnosticFields.locationFor(
            new CellMutationAction.SetCell(formulaCell("SUM(A1:A2)"))));
  }

  @Test
  void selectorLocationsAndWorkbookStepFallbacksStayTyped() {
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange(
            new SheetLabel("Budget"), "BudgetTotal"),
        ExecutionDiagnosticFields.locationFor(
            new NamedRangeSelector.SheetScope(
                new DefinedName("BudgetTotal"), new SheetName("Budget"))));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.NamedRange("BudgetTotal"),
        ExecutionDiagnosticFields.locationFor(
            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal"))));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.namedRangeNameFor(
            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal"))));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Cell(new SheetLabel("Budget"), "B4"),
        ExecutionDiagnosticFields.locationFor(
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("B4"))));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Range(new SheetLabel("Budget"), "A1:B2"),
        ExecutionDiagnosticFields.locationFor(
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2"))));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Sheet(new SheetLabel("Budget")),
        ExecutionDiagnosticFields.locationFor(new SheetSelector.ByName(new SheetName("Budget"))));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.unknown(),
        ExecutionDiagnosticFields.locationFor(new WorkbookSelector.Current()));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.unknown(),
        ExecutionDiagnosticFields.locationFor(new PresenceAssertion.TablePresent()));

    MutationStep pivotAddressStep =
        new MutationStep(
            new StepId("set-pivot"),
            new PivotTableSelector.ByNameOnSheet("SalesPivot", new SheetName("Budget")),
            new StructuredMutationAction.SetPivotTable(
                new PivotTableInput(
                    "SalesPivot",
                    new SheetName("Budget"),
                    new PivotTableInput.PivotSourceSpec.Range(
                        new SheetName("Budget"), new RangeRef("A1:B5")),
                    new PivotTableInput.Anchor(new CellAddress("C5")),
                    List.of(),
                    List.of(),
                    List.of(),
                    List.of(
                        new PivotTableInput.DataField(
                            "Amount",
                            dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction
                                .SUM,
                            "Amount",
                            Optional.empty())))));
    MutationStep pivotNamedRangeStep =
        new MutationStep(
            new StepId("set-pivot-named-range"),
            new PivotTableSelector.ByNameOnSheet("SalesPivot", new SheetName("Budget")),
            new StructuredMutationAction.SetPivotTable(
                new PivotTableInput(
                    "SalesPivot",
                    new SheetName("Budget"),
                    new PivotTableInput.PivotSourceSpec.NamedRange(new DefinedName("BudgetSource")),
                    new PivotTableInput.Anchor(new CellAddress("C5")),
                    List.of(),
                    List.of(),
                    List.of(),
                    List.of(
                        new PivotTableInput.DataField(
                            "Amount",
                            dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction
                                .SUM,
                            "Amount",
                            Optional.empty())))));
    MutationStep tableStep =
        new MutationStep(
            new StepId("set-table"),
            new TableSelector.ByNameOnSheet(new TableName("BudgetTable"), new SheetName("Budget")),
            new StructuredMutationAction.SetTable(
                TableInput.withDefaultMetadata(
                    new TableName("BudgetTable"),
                    new SheetName("Budget"),
                    new RangeRef("A1:B4"),
                    false,
                    new TableStyleInput.None())));
    MutationStep formulaNamedRangeStep =
        new MutationStep(
            new StepId("set-name"),
            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetExpr")),
            new StructuredMutationAction.SetNamedRange(
                new DefinedName("BudgetExpr"),
                new NamedRangeScope.Workbook(),
                NamedRangeTarget.formula("SUM(Budget!B2:B4)")));
    AssertionStep formulaAssertionStep =
        new AssertionStep(
            new StepId("assert-formula"),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("B4")),
            new CellAssertion.FormulaText("SUM(A1:A2)"));
    InspectionStep workbookInspection =
        new InspectionStep(
            new StepId("summary"),
            new WorkbookSelector.Current(),
            new WorkbookIntrospectionQuery.GetWorkbookSummary());

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
        ExecutionExceptionDiagnosticFields.rangeFor(
            new InvalidRangeAddressException("A1:B2", null)));
    assertEquals(
        Optional.of("SUM("),
        ExecutionExceptionDiagnosticFields.formulaFor(
            new InvalidFormulaException("Budget", "B4", "SUM(", "bad", null)));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionExceptionDiagnosticFields.namedRangeNameFor(
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertInstanceOf(
        ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell.class,
        ExecutionDiagnosticFields.locationFor(
            new MutationStep(
                new StepId("set-formula"),
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("B4")),
                new CellMutationAction.SetCell(formulaCell("SUM(A1:A2)"))),
            new InvalidFormulaException("Budget", "B4", "SUM(A1:A2)", "bad", null)));
  }
}
