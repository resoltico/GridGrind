package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
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
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused current-API coverage for executor-owned helper and conversion seams. */
class DefaultGridGrindRequestExecutorCoverageTest {
  @Test
  void classifiesExceptionsAndEnrichesExecuteStepContexts() {
    assertEquals(
        GridGrindProblemCode.WORKBOOK_NOT_FOUND,
        ExecutionResponseSupport.problemCodeFor(
            new WorkbookNotFoundException(Path.of("/tmp/missing.xlsx"))));

    GridGrindResponse.ProblemContext.ExecuteStep context =
        new GridGrindResponse.ProblemContext.ExecuteStep(
            "NEW", "NONE", 0, "step-01-set-cell", "SET_CELL", null, null, null, null, null, null);

    InvalidFormulaException invalidFormula =
        new InvalidFormulaException("Budget", "B4", "SUM(", "invalid", null);
    GridGrindResponse.ProblemContext.ExecuteStep enrichedFormula =
        assertInstanceOf(
            GridGrindResponse.ProblemContext.ExecuteStep.class,
            ExecutionResponseSupport.enrichContext(context, invalidFormula));
    assertEquals("Budget", enrichedFormula.sheetName());
    assertEquals("B4", enrichedFormula.address());
    assertEquals("SUM(", enrichedFormula.formula());

    GridGrindResponse.ProblemContext.ExecuteStep enrichedRange =
        assertInstanceOf(
            GridGrindResponse.ProblemContext.ExecuteStep.class,
            ExecutionResponseSupport.enrichContext(
                context, new InvalidRangeAddressException("A1:B2", null)));
    assertEquals("A1:B2", enrichedRange.range());

    GridGrindResponse.ProblemContext.ExecuteStep enrichedNamedRange =
        assertInstanceOf(
            GridGrindResponse.ProblemContext.ExecuteStep.class,
            ExecutionResponseSupport.enrichContext(
                context,
                new NamedRangeNotFoundException(
                    "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertEquals("BudgetTotal", enrichedNamedRange.namedRangeName());
  }

  @Test
  void extractsSheetAddressRangeFormulaAndNamedRangeMetadataFromSteps() {
    MutationStep setCell =
        new MutationStep(
            "step-01-set-cell",
            new CellSelector.ByAddress("Budget", "B4"),
            new MutationAction.SetCell(formulaCell("SUM(B2:B3)")));
    MutationStep setNamedRange =
        new MutationStep(
            "step-02-set-named-range",
            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
            new MutationAction.SetNamedRange(
                "BudgetTotal",
                new NamedRangeScope.Workbook(),
                new NamedRangeTarget("Budget", "B4")));
    InspectionStep getCells =
        new InspectionStep(
            "step-03-get-cells",
            new CellSelector.ByAddresses("Budget", List.of("B4")),
            new InspectionQuery.GetCells());
    InspectionStep getNamedRanges =
        new InspectionStep(
            "step-04-get-named-ranges",
            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
            new InspectionQuery.GetNamedRanges());

    assertEquals("Budget", ExecutionDiagnosticFields.sheetNameFor(setCell));
    assertEquals("B4", ExecutionDiagnosticFields.addressFor(setCell));
    assertEquals("SUM(B2:B3)", ExecutionDiagnosticFields.formulaFor(setCell));
    assertNull(ExecutionDiagnosticFields.rangeFor(setCell));

    assertEquals("BudgetTotal", ExecutionDiagnosticFields.namedRangeNameFor(setNamedRange));
    assertEquals("Budget", ExecutionDiagnosticFields.sheetNameFor(getCells));
    assertNull(ExecutionDiagnosticFields.addressFor(getCells));
    assertEquals("BudgetTotal", ExecutionDiagnosticFields.namedRangeNameFor(getNamedRanges));

    InvalidFormulaException invalidFormula =
        new InvalidFormulaException("Budget", "C7", "SUM(", "invalid", null);
    assertEquals("Budget", ExecutionDiagnosticFields.sheetNameFor(getNamedRanges, invalidFormula));
    assertEquals("C7", ExecutionDiagnosticFields.addressFor(getNamedRanges, invalidFormula));
    assertEquals("SUM(", ExecutionDiagnosticFields.formulaFor(getNamedRanges, invalidFormula));
    assertEquals(
        "A1:B9",
        ExecutionDiagnosticFields.rangeFor(
            new InspectionStep(
                "step-05-get-window",
                new RangeSelector.RectangularWindow("Budget", "A1", 9, 2),
                new InspectionQuery.GetSheetSchema()),
            invalidFormula));
  }

  @Test
  void convertsStepNativeMutationsAndInspectionsIntoWorkbookCoreCommands() {
    MutationStep setNamedRange =
        new MutationStep(
            "step-01-set-named-range",
            new NamedRangeSelector.WorkbookScope("BudgetExpr"),
            new MutationAction.SetNamedRange(
                "BudgetExpr",
                new NamedRangeScope.Workbook(),
                new NamedRangeTarget("SUM(Budget!A1:A3)")));
    InspectionStep getNamedRanges =
        new InspectionStep(
            "step-03-get-named-ranges",
            new NamedRangeSelector.WorkbookScope("BudgetExpr"),
            new InspectionQuery.GetNamedRanges());
    InspectionStep workbookFindings =
        new InspectionStep(
            "step-04-workbook-findings",
            new WorkbookSelector.Current(),
            new InspectionQuery.AnalyzeWorkbookFindings());

    WorkbookCommand.SetNamedRange namedRangeCommand =
        assertInstanceOf(
            WorkbookCommand.SetNamedRange.class, WorkbookCommandConverter.toCommand(setNamedRange));
    WorkbookReadCommand.GetNamedRanges namedRangesCommand =
        assertInstanceOf(
            WorkbookReadCommand.GetNamedRanges.class,
            InspectionCommandConverter.toReadCommand(getNamedRanges));
    assertInstanceOf(
        WorkbookReadCommand.AnalyzeWorkbookFindings.class,
        InspectionCommandConverter.toReadCommand(workbookFindings));

    assertEquals("BudgetExpr", namedRangeCommand.definition().name());
    assertEquals("step-03-get-named-ranges", namedRangesCommand.stepId());
  }

  @Test
  void convertsFocusedWorkbookReadResultsIntoInspectionResults() {
    InspectionResult.WorkbookProtectionResult protection =
        assertInstanceOf(
            InspectionResult.WorkbookProtectionResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookProtectionResult(
                    "step-01-protection",
                    new ExcelWorkbookProtectionSnapshot(true, false, true, true, false))));
    InspectionResult.CellsResult cells =
        assertInstanceOf(
            InspectionResult.CellsResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.CellsResult(
                    "step-02-cells",
                    "Budget",
                    List.of(
                        new ExcelCellSnapshot.NumberSnapshot(
                            "B4",
                            "NUMBER",
                            "61",
                            defaultStyle(),
                            ExcelCellMetadataSnapshot.empty(),
                            61.0d)))));

    assertTrue(protection.protection().structureLocked());
    assertEquals("step-02-cells", cells.stepId());
    assertEquals("Budget", cells.sheetName());
    assertEquals("B4", cells.cells().getFirst().address());
  }

  private static ExcelCellStyleSnapshot defaultStyle() {
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
        new ExcelCellFillSnapshot(ExcelFillPattern.NONE, null, null),
        new ExcelBorderSnapshot(
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null)),
        new ExcelCellProtectionSnapshot(true, false));
  }
}
