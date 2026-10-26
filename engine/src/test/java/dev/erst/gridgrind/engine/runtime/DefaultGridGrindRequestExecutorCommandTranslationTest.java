package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Workbook command and read-command translation coverage. */
class DefaultGridGrindRequestExecutorCommandTranslationTest
    extends DefaultGridGrindRequestExecutorTestSupport {

  @Test
  void convertsStructuralAndSurfaceReadOperationsIntoWorkbookReadCommands() {
    WorkbookReadCommand workbookSummary =
        readCommand(
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary()));
    WorkbookReadCommand workbookProtection =
        readCommand(
            inspect(
                "workbook-protection",
                new WorkbookSelector.Current(),
                new WorkbookIntrospectionQuery.GetWorkbookProtection()));
    WorkbookReadCommand namedRanges =
        readCommand(
            inspect(
                "ranges",
                new dev.erst.gridgrind.core.plan.NamedRangeSelector.AnyOf(
                    List.of(
                        new dev.erst.gridgrind.core.plan.NamedRangeSelector.ByName(
                            new DefinedName("BudgetTotal")),
                        new dev.erst.gridgrind.core.plan.NamedRangeSelector.SheetScope(
                            new DefinedName("LocalItem"), new SheetName("Budget")))),
                new WorkbookIntrospectionQuery.GetNamedRanges()));
    WorkbookReadCommand sheetSummary =
        readCommand(
            inspect(
                "sheet",
                new SheetSelector.ByName(new SheetName("Budget")),
                new SheetIntrospectionQuery.GetSheetSummary()));
    WorkbookReadCommand cells =
        readCommand(
            inspect(
                "cells",
                new CellSelector.ByAddresses(
                    new SheetName("Budget"), List.of(new CellAddress("A1"))),
                new SheetIntrospectionQuery.GetCells()));
    WorkbookReadCommand window =
        readCommand(
            inspect(
                "window",
                new RangeSelector.RectangularWindow(
                    new SheetName("Budget"), new CellAddress("A1"), 2, 2),
                new SheetIntrospectionQuery.GetWindow()));
    WorkbookReadCommand merged =
        readCommand(
            inspect(
                "merged",
                new SheetSelector.ByName(new SheetName("Budget")),
                new SheetIntrospectionQuery.GetMergedRegions()));
    WorkbookReadCommand hyperlinks =
        readCommand(
            inspect(
                "hyperlinks",
                new CellSelector.AllUsedInSheet(new SheetName("Budget")),
                new SheetIntrospectionQuery.GetHyperlinks()));
    WorkbookReadCommand comments =
        readCommand(
            inspect(
                "comments",
                new CellSelector.ByAddresses(
                    new SheetName("Budget"), List.of(new CellAddress("A1"))),
                new SheetIntrospectionQuery.GetComments()));
    WorkbookReadCommand layout =
        readCommand(
            inspect(
                "layout",
                new SheetSelector.ByName(new SheetName("Budget")),
                new SheetIntrospectionQuery.GetSheetLayout()));
    WorkbookReadCommand printLayout =
        readCommand(
            inspect(
                "printLayout",
                new SheetSelector.ByName(new SheetName("Budget")),
                new SheetIntrospectionQuery.GetPrintLayout()));
    WorkbookReadCommand validations =
        readCommand(
            inspect(
                "validations",
                new RangeSelector.AllOnSheet(new SheetName("Budget")),
                new SheetIntrospectionQuery.GetDataValidations()));
    WorkbookReadCommand conditionalFormatting =
        readCommand(
            inspect(
                "conditionalFormatting",
                new RangeSelector.ByRanges(new SheetName("Budget"), List.of(new RangeRef("A1:A3"))),
                new SheetIntrospectionQuery.GetConditionalFormatting()));
    WorkbookReadCommand autofilters =
        readCommand(
            inspect(
                "autofilters",
                new SheetSelector.ByName(new SheetName("Budget")),
                new SheetIntrospectionQuery.GetAutofilters()));
    WorkbookReadCommand tables =
        readCommand(
            inspect(
                "tables",
                new TableSelector.ByNames(List.of(new TableName("BudgetTable"))),
                new WorkbookAssetIntrospectionQuery.GetTables()));
    WorkbookReadCommand formulaSurface =
        readCommand(
            inspect(
                "formula",
                new SheetSelector.ByNames(List.of(new SheetName("Budget"))),
                new InspectionSurfaceQuery.GetFormulaSurface()));
    WorkbookReadCommand schema =
        readCommand(
            inspect(
                "schema",
                new RangeSelector.RectangularWindow(
                    new SheetName("Budget"), new CellAddress("A1"), 3, 2),
                new InspectionSurfaceQuery.GetSheetSchema()));
    WorkbookReadCommand namedRangeSurface =
        readCommand(
            inspect(
                "surface",
                new dev.erst.gridgrind.core.plan.NamedRangeSelector.All(),
                new InspectionSurfaceQuery.GetNamedRangeSurface()));

    assertInstanceOf(WorkbookReadCommand.GetWorkbookSummary.class, workbookSummary);
    assertInstanceOf(WorkbookReadCommand.GetWorkbookProtection.class, workbookProtection);
    assertInstanceOf(WorkbookReadCommand.GetNamedRanges.class, namedRanges);
    assertInstanceOf(WorkbookReadCommand.GetSheetSummary.class, sheetSummary);
    assertInstanceOf(WorkbookReadCommand.GetCells.class, cells);
    assertInstanceOf(WorkbookReadCommand.GetWindow.class, window);
    assertInstanceOf(WorkbookReadCommand.GetMergedRegions.class, merged);
    assertInstanceOf(WorkbookReadCommand.GetHyperlinks.class, hyperlinks);
    assertInstanceOf(WorkbookReadCommand.GetComments.class, comments);
    assertInstanceOf(WorkbookReadCommand.GetSheetLayout.class, layout);
    assertInstanceOf(WorkbookReadCommand.GetPrintLayout.class, printLayout);
    assertInstanceOf(WorkbookReadCommand.GetDataValidations.class, validations);
    assertInstanceOf(WorkbookReadCommand.GetConditionalFormatting.class, conditionalFormatting);
    assertInstanceOf(WorkbookReadCommand.GetAutofilters.class, autofilters);
    assertInstanceOf(WorkbookReadCommand.GetTables.class, tables);
    assertInstanceOf(WorkbookReadCommand.GetFormulaSurface.class, formulaSurface);
    assertInstanceOf(WorkbookReadCommand.GetSheetSchema.class, schema);
    assertInstanceOf(WorkbookReadCommand.GetNamedRangeSurface.class, namedRangeSurface);
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelNamedRangeSelection.Selected.class,
        cast(WorkbookReadCommand.GetNamedRanges.class, namedRanges).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelCellSelection.AllUsedCells.class,
        cast(WorkbookReadCommand.GetHyperlinks.class, hyperlinks).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelCellSelection.Selected.class,
        cast(WorkbookReadCommand.GetComments.class, comments).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelRangeSelection.All.class,
        cast(WorkbookReadCommand.GetDataValidations.class, validations).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelRangeSelection.Selected.class,
        cast(WorkbookReadCommand.GetConditionalFormatting.class, conditionalFormatting)
            .selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelTableSelection.ByNames.class,
        cast(WorkbookReadCommand.GetTables.class, tables).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelSheetSelection.Selected.class,
        cast(WorkbookReadCommand.GetFormulaSurface.class, formulaSurface).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelNamedRangeSelection.All.class,
        cast(WorkbookReadCommand.GetNamedRangeSurface.class, namedRangeSurface).selection());
  }

  @Test
  void convertsAnalysisReadOperationsIntoWorkbookReadCommands() {
    WorkbookReadCommand formulaHealth =
        readCommand(
            inspect(
                "formulaHealth",
                new SheetSelector.All(),
                new InspectionAnalysisQuery.AnalyzeFormulaHealth()));
    WorkbookReadCommand validationHealth =
        readCommand(
            inspect(
                "validationHealth",
                new SheetSelector.All(),
                new InspectionAnalysisQuery.AnalyzeDataValidationHealth()));
    WorkbookReadCommand conditionalFormattingHealth =
        readCommand(
            inspect(
                "conditionalFormattingHealth",
                new SheetSelector.ByNames(List.of(new SheetName("Budget"))),
                new InspectionAnalysisQuery.AnalyzeConditionalFormattingHealth()));
    WorkbookReadCommand autofilterHealth =
        readCommand(
            inspect(
                "autofilterHealth",
                new SheetSelector.ByNames(List.of(new SheetName("Budget"))),
                new InspectionAnalysisQuery.AnalyzeAutofilterHealth()));
    WorkbookReadCommand tableHealth =
        readCommand(
            inspect(
                "tableHealth",
                new TableSelector.ByNames(List.of(new TableName("BudgetTable"))),
                new InspectionAnalysisQuery.AnalyzeTableHealth()));
    WorkbookReadCommand hyperlinkHealth =
        readCommand(
            inspect(
                "hyperlinkHealth",
                new SheetSelector.All(),
                new InspectionAnalysisQuery.AnalyzeHyperlinkHealth()));
    WorkbookReadCommand namedRangeHealth =
        readCommand(
            inspect(
                "namedRangeHealth",
                new dev.erst.gridgrind.core.plan.NamedRangeSelector.All(),
                new InspectionAnalysisQuery.AnalyzeNamedRangeHealth()));
    WorkbookReadCommand workbookFindings =
        readCommand(
            inspect(
                "workbookFindings",
                new WorkbookSelector.Current(),
                new InspectionAnalysisQuery.AnalyzeWorkbookFindings()));

    assertInstanceOf(WorkbookReadCommand.AnalyzeFormulaHealth.class, formulaHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeDataValidationHealth.class, validationHealth);
    assertInstanceOf(
        WorkbookReadCommand.AnalyzeConditionalFormattingHealth.class, conditionalFormattingHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeAutofilterHealth.class, autofilterHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeTableHealth.class, tableHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeHyperlinkHealth.class, hyperlinkHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeNamedRangeHealth.class, namedRangeHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeWorkbookFindings.class, workbookFindings);
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelSheetSelection.Selected.class,
        cast(
                WorkbookReadCommand.AnalyzeConditionalFormattingHealth.class,
                conditionalFormattingHealth)
            .selection());
  }
}
