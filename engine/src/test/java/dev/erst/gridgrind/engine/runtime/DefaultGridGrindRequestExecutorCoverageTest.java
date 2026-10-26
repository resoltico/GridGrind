package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.*;
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
                    "BudgetTotal", new NamedRangeScopeReport.Workbook())));
    assertEquals(java.util.Optional.of("BudgetTotal"), enrichedNamedRange.namedRangeName());
  }

  @Test
  void extractsSheetAddressRangeFormulaAndNamedRangeMetadataFromSteps() {
    MutationStep setCell =
        new MutationStep(
            new StepId("step-01-set-cell"),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("B4")),
            new CellMutationAction.SetCell(formulaCell("SUM(B2:B3)")));
    MutationStep setNamedRange =
        new MutationStep(
            new StepId("step-02-set-named-range"),
            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal")),
            new StructuredMutationAction.SetNamedRange(
                new DefinedName("BudgetTotal"),
                new NamedRangeScope.Workbook(),
                NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("B4"))));
    InspectionStep getCells =
        new InspectionStep(
            new StepId("step-03-get-cells"),
            new CellSelector.ByAddresses(new SheetName("Budget"), List.of(new CellAddress("B4"))),
            new SheetIntrospectionQuery.GetCells());
    InspectionStep getNamedRanges =
        new InspectionStep(
            new StepId("step-04-get-named-ranges"),
            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal")),
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
                new RangeSelector.RectangularWindow(
                    new SheetName("Budget"), new CellAddress("A1"), 9, 2),
                new InspectionSurfaceQuery.GetSheetSchema()),
            invalidFormula));
  }

  @Test
  void convertsStepNativeMutationsAndInspectionsIntoWorkbookCoreCommands() {
    InspectionStep getNamedRanges =
        new InspectionStep(
            new StepId("step-03-get-named-ranges"),
            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetExpr")),
            new WorkbookIntrospectionQuery.GetNamedRanges());
    InspectionStep workbookFindings =
        new InspectionStep(
            new StepId("step-04-workbook-findings"),
            new WorkbookSelector.Current(),
            new InspectionAnalysisQuery.AnalyzeWorkbookFindings());

    WorkbookReadCommand.GetNamedRanges namedRangesCommand =
        assertInstanceOf(
            WorkbookReadCommand.GetNamedRanges.class,
            InspectionCommandConverter.toReadCommand(getNamedRanges));
    assertInstanceOf(
        WorkbookReadCommand.AnalyzeWorkbookFindings.class,
        InspectionCommandConverter.toReadCommand(workbookFindings));

    assertEquals("step-03-get-named-ranges", namedRangesCommand.stepId().value());
  }

  static CellStyleReport defaultStyle() {
    return new CellStyleReport(
        "",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Aptos",
            FontHeightReport.fromPoints(new BigDecimal("11")),
            null,
            false,
            false),
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None()),
        new CellProtectionReport(true, false));
  }
}
