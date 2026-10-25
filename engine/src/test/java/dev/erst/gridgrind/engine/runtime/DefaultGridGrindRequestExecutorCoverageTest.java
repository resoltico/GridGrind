package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelBorderSideSnapshot;
import dev.erst.gridgrind.excel.ExcelBorderSnapshot;
import dev.erst.gridgrind.excel.ExcelCellAlignmentSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFillSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelCellMetadataSnapshot;
import dev.erst.gridgrind.excel.ExcelCellProtectionSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSnapshot;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused current-API coverage for executor-owned helper and conversion seams. */
class DefaultGridGrindRequestExecutorCoverageTest
    extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void classifiesExceptionsAndEnrichesExecuteStepContexts() {
    assertEquals(
        GridGrindProblemCode.WORKBOOK_NOT_FOUND,
        ExecutionResponseSupport.problemCodeFor(
            new WorkbookNotFoundException(Path.of("/tmp/missing.xlsx"))));

    dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep context =
        new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep(
            dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape.known(
                "NEW", "NONE"),
            new dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.StepReference(
                0, new StepId("step-01-set-cell"), "MUTATION", "SET_CELL"),
            dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.ProblemLocation
                .unknown());

    InvalidFormulaException invalidFormula =
        new InvalidFormulaException("Budget", "B4", "SUM(", "invalid", null);
    dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep enrichedFormula =
        assertInstanceOf(
            dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep.class,
            ExecutionResponseSupport.enrichContext(context, invalidFormula));
    assertEquals(
        java.util.Optional.of("Budget"),
        enrichedFormula.sheetName().map(dev.erst.gridgrind.core.vocabulary.WireString::value));
    assertEquals(java.util.Optional.of("B4"), enrichedFormula.address());
    assertEquals(java.util.Optional.of("SUM("), enrichedFormula.formula());

    dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep enrichedRange =
        assertInstanceOf(
            dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep.class,
            ExecutionResponseSupport.enrichContext(
                context, new InvalidRangeAddressException("A1:B2", null)));
    assertEquals(java.util.Optional.of("A1:B2"), enrichedRange.range());

    dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep enrichedNamedRange =
        assertInstanceOf(
            dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep.class,
            ExecutionResponseSupport.enrichContext(
                context,
                new NamedRangeNotFoundException(
                    "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertEquals(java.util.Optional.of("BudgetTotal"), enrichedNamedRange.namedRangeName());
  }

  @Test
  void extractsSheetAddressRangeFormulaAndNamedRangeMetadataFromSteps() {
    MutationStep setCell =
        new MutationStep(
            new StepId("step-01-set-cell"),
            new CellSelector.ByAddress(new SheetName("Budget"), "B4"),
            new CellMutationAction.SetCell(formulaCell("SUM(B2:B3)")));
    MutationStep setNamedRange =
        new MutationStep(
            new StepId("step-02-set-named-range"),
            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
            new StructuredMutationAction.SetNamedRange(
                "BudgetTotal",
                new NamedRangeScope.Workbook(),
                NamedRangeTarget.range(new SheetName("Budget"), "B4")));
    InspectionStep getCells =
        new InspectionStep(
            new StepId("step-03-get-cells"),
            new CellSelector.ByAddresses(new SheetName("Budget"), List.of("B4")),
            new SheetIntrospectionQuery.GetCells());
    InspectionStep getNamedRanges =
        new InspectionStep(
            new StepId("step-04-get-named-ranges"),
            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
            new WorkbookIntrospectionQuery.GetNamedRanges());

    assertEquals(java.util.Optional.of("Budget"), ExecutionDiagnosticFields.sheetNameFor(setCell));
    assertEquals(java.util.Optional.of("B4"), ExecutionDiagnosticFields.addressFor(setCell));
    assertEquals(
        java.util.Optional.of("SUM(B2:B3)"), ExecutionDiagnosticFields.formulaFor(setCell));
    assertEquals(java.util.Optional.empty(), ExecutionDiagnosticFields.rangeFor(setCell));

    assertEquals(
        java.util.Optional.of("BudgetTotal"),
        ExecutionDiagnosticFields.namedRangeNameFor(setNamedRange));
    assertEquals(java.util.Optional.of("Budget"), ExecutionDiagnosticFields.sheetNameFor(getCells));
    assertEquals(java.util.Optional.empty(), ExecutionDiagnosticFields.addressFor(getCells));
    assertEquals(
        java.util.Optional.of("BudgetTotal"),
        ExecutionDiagnosticFields.namedRangeNameFor(getNamedRanges));

    InvalidFormulaException invalidFormula =
        new InvalidFormulaException("Budget", "C7", "SUM(", "invalid", null);
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionDiagnosticFields.sheetNameFor(getNamedRanges, invalidFormula));
    assertEquals(
        java.util.Optional.of("C7"),
        ExecutionDiagnosticFields.addressFor(getNamedRanges, invalidFormula));
    assertEquals(
        java.util.Optional.of("SUM("),
        ExecutionDiagnosticFields.formulaFor(getNamedRanges, invalidFormula));
    assertEquals(
        java.util.Optional.of("A1:B9"),
        ExecutionDiagnosticFields.rangeFor(
            new InspectionStep(
                new StepId("step-05-get-window"),
                new RangeSelector.RectangularWindow(new SheetName("Budget"), "A1", 9, 2),
                new InspectionSurfaceQuery.GetSheetSchema()),
            invalidFormula));
  }

  @Test
  void convertsStepNativeMutationsAndInspectionsIntoWorkbookCoreCommands() {
    MutationStep setNamedRange =
        new MutationStep(
            new StepId("step-01-set-named-range"),
            new NamedRangeSelector.WorkbookScope("BudgetExpr"),
            new StructuredMutationAction.SetNamedRange(
                "BudgetExpr",
                new NamedRangeScope.Workbook(),
                NamedRangeTarget.formula("SUM(Budget!A1:A3)")));
    InspectionStep getNamedRanges =
        new InspectionStep(
            new StepId("step-03-get-named-ranges"),
            new NamedRangeSelector.WorkbookScope("BudgetExpr"),
            new WorkbookIntrospectionQuery.GetNamedRanges());
    InspectionStep workbookFindings =
        new InspectionStep(
            new StepId("step-04-workbook-findings"),
            new WorkbookSelector.Current(),
            new InspectionAnalysisQuery.AnalyzeWorkbookFindings());

    WorkbookMetadataCommand.SetNamedRange namedRangeCommand =
        assertInstanceOf(
            WorkbookMetadataCommand.SetNamedRange.class,
            WorkbookCommandConverter.toCommand(setNamedRange));
    WorkbookReadCommand.GetNamedRanges namedRangesCommand =
        assertInstanceOf(
            WorkbookReadCommand.GetNamedRanges.class,
            InspectionCommandConverter.toReadCommand(getNamedRanges));
    assertInstanceOf(
        WorkbookReadCommand.AnalyzeWorkbookFindings.class,
        InspectionCommandConverter.toReadCommand(workbookFindings));

    assertEquals("BudgetExpr", namedRangeCommand.definition().name());
    assertEquals("step-03-get-named-ranges", namedRangesCommand.stepId().value());
  }

  @Test
  void convertsFocusedWorkbookReadResultsIntoInspectionResults() {
    WorkbookInspectionResult.WorkbookProtectionResult protection =
        assertInstanceOf(
            WorkbookInspectionResult.WorkbookProtectionResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookProtectionResult(
                    new StepId("step-01-protection"),
                    new ExcelWorkbookProtectionSnapshot(true, false, true, true, false))));
    SheetInspectionResult.CellsResult cells =
        assertInstanceOf(
            SheetInspectionResult.CellsResult.class,
            InspectionResultConverter.toReadResult(
                excelCellsResult(
                    "step-02-cells",
                    "Budget",
                    List.of(
                        new ExcelCellSnapshot.NumberSnapshot(
                            "B4",
                            "61",
                            defaultStyle(),
                            ExcelCellMetadataSnapshot.empty(),
                            61.0d)))));

    assertTrue(protection.protection().structureLocked());
    assertEquals("step-02-cells", cells.stepId().value());
    assertEquals("Budget", cells.sheetName().value());
    assertEquals("B4", cells.cells().getFirst().address());
  }

  static ExcelCellStyleSnapshot defaultStyle() {
    return new ExcelCellStyleSnapshot(
        "",
        new ExcelCellAlignmentSnapshot(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new ExcelCellFontSnapshot(
            false,
            false,
            "Aptos",
            ExcelFontHeight.fromPoints(new BigDecimal("11")),
            null,
            false,
            false),
        ExcelCellFillSnapshot.pattern(ExcelFillPattern.NONE),
        new ExcelBorderSnapshot(
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null)),
        new ExcelCellProtectionSnapshot(true, false));
  }
}
