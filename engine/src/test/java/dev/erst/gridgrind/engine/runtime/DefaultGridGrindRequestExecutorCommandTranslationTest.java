package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.layout.PaneSpec;
import dev.erst.gridgrind.core.model.layout.SheetDefaultsSpec;
import dev.erst.gridgrind.core.model.layout.SheetDisplaySpec;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummarySpec;
import dev.erst.gridgrind.core.model.layout.SheetPresentationSpec;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintLayoutSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorSpec;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Workbook command and read-command translation coverage. */
class DefaultGridGrindRequestExecutorCommandTranslationTest
    extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void convertsWaveThreeOperationsIntoWorkbookCommands() {
    WorkbookCommand setHyperlink =
        command(
            mutate(
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                new CellMutationAction.SetHyperlink(
                    new HyperlinkTarget.Url("https://example.com/report"))));
    WorkbookCommand clearHyperlink =
        command(
            mutate(
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                new CellMutationAction.ClearHyperlink()));
    WorkbookCommand setComment =
        command(
            mutate(
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                new CellMutationAction.SetComment(
                    CommentSpec.plain(text("Review"), "GridGrind", false))));
    WorkbookCommand clearComment =
        command(
            mutate(
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                new CellMutationAction.ClearComment()));
    WorkbookCommand setNamedRange =
        command(
            ExecutorStructuredMutationTargets.namedRange(
                new StructuredMutationAction.SetNamedRange(
                    new DefinedName("BudgetTotal"),
                    new NamedRangeScope.Workbook(),
                    NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("B4")))));
    WorkbookCommand deleteNamedRange =
        command(
            mutate(
                new dev.erst.gridgrind.core.plan.NamedRangeSelector.SheetScope(
                    new DefinedName("BudgetTotal"), new SheetName("Budget")),
                new StructuredMutationAction.DeleteNamedRange()));

    assertInstanceOf(WorkbookAnnotationCommand.SetHyperlink.class, setHyperlink);
    assertInstanceOf(WorkbookAnnotationCommand.ClearHyperlink.class, clearHyperlink);
    assertInstanceOf(WorkbookAnnotationCommand.SetComment.class, setComment);
    assertInstanceOf(WorkbookAnnotationCommand.ClearComment.class, clearComment);
    assertInstanceOf(WorkbookMetadataCommand.SetNamedRange.class, setNamedRange);
    assertInstanceOf(WorkbookMetadataCommand.DeleteNamedRange.class, deleteNamedRange);
    assertEquals(
        new HyperlinkTarget.Url("https://example.com/report"),
        cast(WorkbookAnnotationCommand.SetHyperlink.class, setHyperlink).target());
    assertEquals(
        new ResolvedComment("Review", "GridGrind", false),
        cast(WorkbookAnnotationCommand.SetComment.class, setComment).comment());
    assertEquals(
        new ExcelNamedRangeDefinition(
            "BudgetTotal",
            new ExcelNamedRangeScope.WorkbookScope(),
            ExcelNamedRangeTarget.range("Budget", "B4")),
        cast(WorkbookMetadataCommand.SetNamedRange.class, setNamedRange).definition());
    assertEquals(
        new ExcelNamedRangeScope.SheetScope("Budget"),
        cast(WorkbookMetadataCommand.DeleteNamedRange.class, deleteNamedRange).scope());
  }

  @Test
  void convertsRemainingWorkbookOperationsIntoWorkbookCommands() {
    assertInstanceOf(
        WorkbookSheetCommand.CreateSheet.class,
        command(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.EnsureSheet())));
    assertInstanceOf(
        WorkbookSheetCommand.RenameSheet.class,
        command(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.RenameSheet(new SheetName("Summary")))));
    assertInstanceOf(
        WorkbookSheetCommand.DeleteSheet.class,
        command(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.DeleteSheet())));
    assertInstanceOf(
        WorkbookSheetCommand.MoveSheet.class,
        command(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.MoveSheet(0))));
    assertInstanceOf(
        WorkbookStructureCommand.MergeCells.class,
        command(
            mutate(
                new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2")),
                new WorkbookMutationAction.MergeCells())));
    assertInstanceOf(
        WorkbookStructureCommand.UnmergeCells.class,
        command(
            mutate(
                new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2")),
                new WorkbookMutationAction.UnmergeCells())));
    assertInstanceOf(
        WorkbookStructureCommand.SetColumnWidth.class,
        command(
            mutate(
                new ColumnBandSelector.Span(new SheetName("Budget"), 0, 1),
                new WorkbookMutationAction.SetColumnWidth(16.0))));
    assertInstanceOf(
        WorkbookStructureCommand.SetRowHeight.class,
        command(
            mutate(
                new RowBandSelector.Span(new SheetName("Budget"), 0, 1),
                new WorkbookMutationAction.SetRowHeight(28.5))));
    assertInstanceOf(
        WorkbookStructureCommand.InsertRows.class,
        command(
            mutate(
                new RowBandSelector.Insertion(new SheetName("Budget"), 1, 2),
                new WorkbookMutationAction.InsertRows())));
    assertInstanceOf(
        WorkbookStructureCommand.DeleteRows.class,
        command(
            mutate(
                new RowBandSelector.Span(new SheetName("Budget"), 1, 2),
                new WorkbookMutationAction.DeleteRows())));
    assertInstanceOf(
        WorkbookStructureCommand.ShiftRows.class,
        command(
            mutate(
                new RowBandSelector.Span(new SheetName("Budget"), 1, 2),
                new WorkbookMutationAction.ShiftRows(1))));
    assertInstanceOf(
        WorkbookStructureCommand.InsertColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Insertion(new SheetName("Budget"), 1, 2),
                new WorkbookMutationAction.InsertColumns())));
    assertInstanceOf(
        WorkbookStructureCommand.DeleteColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span(new SheetName("Budget"), 1, 2),
                new WorkbookMutationAction.DeleteColumns())));
    assertInstanceOf(
        WorkbookStructureCommand.ShiftColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span(new SheetName("Budget"), 1, 2),
                new WorkbookMutationAction.ShiftColumns(-1))));
    assertInstanceOf(
        WorkbookStructureCommand.SetRowVisibility.class,
        command(
            mutate(
                new RowBandSelector.Span(new SheetName("Budget"), 1, 2),
                new WorkbookMutationAction.SetRowVisibility(true))));
    assertInstanceOf(
        WorkbookStructureCommand.SetColumnVisibility.class,
        command(
            mutate(
                new ColumnBandSelector.Span(new SheetName("Budget"), 1, 2),
                new WorkbookMutationAction.SetColumnVisibility(true))));
    assertInstanceOf(
        WorkbookStructureCommand.GroupRows.class,
        command(
            mutate(
                new RowBandSelector.Span(new SheetName("Budget"), 1, 2),
                new WorkbookMutationAction.GroupRows(false))));
    assertInstanceOf(
        WorkbookStructureCommand.UngroupRows.class,
        command(
            mutate(
                new RowBandSelector.Span(new SheetName("Budget"), 1, 2),
                new WorkbookMutationAction.UngroupRows())));
    assertInstanceOf(
        WorkbookStructureCommand.GroupColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span(new SheetName("Budget"), 1, 2),
                new WorkbookMutationAction.GroupColumns(false))));
    assertInstanceOf(
        WorkbookStructureCommand.UngroupColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span(new SheetName("Budget"), 1, 2),
                new WorkbookMutationAction.UngroupColumns())));
    assertInstanceOf(
        WorkbookLayoutCommand.SetSheetPane.class,
        command(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.SetSheetPane(new PaneSpec.Frozen(1, 1, 1, 1)))));
    assertInstanceOf(
        WorkbookLayoutCommand.SetSheetZoom.class,
        command(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.SetSheetZoom(125))));
    assertInstanceOf(
        WorkbookLayoutCommand.SetSheetPresentation.class,
        command(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.SetSheetPresentation(
                    new SheetPresentationSpec(
                        new SheetDisplaySpec(false, false, false, true, true),
                        java.util.Optional.of(ColorSpec.rgb(new RgbHex("#112233"))),
                        new SheetOutlineSummarySpec(false, false),
                        new SheetDefaultsSpec(11, 18.5d),
                        List.of(
                            new IgnoredErrorSpec(
                                new RangeRef("A1:B2"),
                                List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))))))));
    assertInstanceOf(
        WorkbookLayoutCommand.SetPrintLayout.class,
        command(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.SetPrintLayout(
                    PrintLayoutSpec.withDefaultSetup(
                        new PrintAreaSpec.Range(new RangeRef("A1:B12")),
                        ExcelPrintOrientation.LANDSCAPE,
                        new PrintScalingSpec.Fit(1, 0),
                        new PrintTitleRowsSpec.Band(0, 0),
                        new PrintTitleColumnsSpec.Band(0, 0),
                        headerFooter("Budget", "", ""),
                        headerFooter("", "Page &P", ""))))));
    assertInstanceOf(
        WorkbookLayoutCommand.ClearPrintLayout.class,
        command(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.ClearPrintLayout())));
    assertInstanceOf(
        WorkbookCellCommand.SetCell.class,
        command(
            mutate(
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                new CellMutationAction.SetCell(textCell("x")))));
    assertInstanceOf(
        WorkbookCellCommand.SetRange.class,
        command(
            mutate(
                new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B1")),
                new CellMutationAction.SetRange(
                    new dev.erst.gridgrind.core.model.value.CellGridSpec.Typed(
                        List.of(List.of(textCell("x"))))))));
    assertInstanceOf(
        WorkbookCellCommand.ClearRange.class,
        command(
            mutate(
                new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B1")),
                new CellMutationAction.ClearRange())));
    assertInstanceOf(
        WorkbookFormattingCommand.SetConditionalFormatting.class,
        command(
            mutate(
                new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:A3")),
                new StructuredMutationAction.SetConditionalFormatting(
                    commandTranslationConditionalFormattingBlock()))));
    assertInstanceOf(
        WorkbookFormattingCommand.ClearConditionalFormatting.class,
        command(
            mutate(
                new RangeSelector.AllOnSheet(new SheetName("Budget")),
                new StructuredMutationAction.ClearConditionalFormatting())));
    assertInstanceOf(
        WorkbookTabularCommand.SetAutofilter.class,
        command(
            mutate(
                new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B4")),
                new StructuredMutationAction.SetAutofilter())));
    assertInstanceOf(
        WorkbookTabularCommand.ClearAutofilter.class,
        command(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new StructuredMutationAction.ClearAutofilter())));
    assertInstanceOf(
        WorkbookTabularCommand.SetTable.class,
        command(
            mutate(
                new StructuredMutationAction.SetTable(
                    TableSpec.withDefaultMetadata(
                        new TableName("BudgetTable"),
                        new SheetName("Budget"),
                        new RangeRef("A1:B4"),
                        false,
                        new TableStyleSpec.None())))));
    assertInstanceOf(
        WorkbookTabularCommand.DeleteTable.class,
        command(
            mutate(
                new TableSelector.ByNameOnSheet(
                    new TableName("BudgetTable"), new SheetName("Budget")),
                new StructuredMutationAction.DeleteTable())));
    assertInstanceOf(
        WorkbookCellCommand.AppendRow.class,
        command(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new CellMutationAction.AppendRow(
                    new dev.erst.gridgrind.core.model.value.CellRowSpec.Typed(
                        List.of(textCell("x")))))));
    assertInstanceOf(
        WorkbookLayoutCommand.AutoSizeColumns.class,
        command(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.AutoSizeColumns())));

    WorkbookLayoutCommand.SetSheetPane setSheetPaneNone =
        cast(
            WorkbookLayoutCommand.SetSheetPane.class,
            command(
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.SetSheetPane(new PaneSpec.None()))));
    WorkbookLayoutCommand.SetSheetPane setSheetPaneSplit =
        cast(
            WorkbookLayoutCommand.SetSheetPane.class,
            command(
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.SetSheetPane(
                        new PaneSpec.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT)))));
    WorkbookLayoutCommand.SetPrintLayout defaultPrintLayout =
        cast(
            WorkbookLayoutCommand.SetPrintLayout.class,
            command(
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.SetPrintLayout(PrintLayoutSpec.defaults()))));
    WorkbookLayoutCommand.SetSheetPresentation defaultSheetPresentation =
        cast(
            WorkbookLayoutCommand.SetSheetPresentation.class,
            command(
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.SetSheetPresentation(
                        SheetPresentationSpec.defaults()))));

    assertEquals(new ExcelSheetPane.None(), setSheetPaneNone.pane());
    assertEquals(
        new ExcelSheetPane.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT),
        setSheetPaneSplit.pane());
    assertEquals(ExcelSheetPresentation.defaults(), defaultSheetPresentation.presentation());
    assertEquals(
        new dev.erst.gridgrind.core.model.print.ResolvedPrintLayout(
            new dev.erst.gridgrind.core.model.print.PrintAreaSpec.None(),
            ExcelPrintOrientation.PORTRAIT,
            new dev.erst.gridgrind.core.model.print.PrintScalingSpec.Automatic(),
            new dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec.None(),
            new dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec.None(),
            new dev.erst.gridgrind.core.model.print.HeaderFooterTextReport("", "", ""),
            new dev.erst.gridgrind.core.model.print.HeaderFooterTextReport("", "", "")),
        defaultPrintLayout.printLayout());
  }

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

  @Test
  void convertsSheetStateOperationsIntoWorkbookCommandsWithExactFields() {
    WorkbookSheetCommand.CopySheet copySheet =
        cast(
            WorkbookSheetCommand.CopySheet.class,
            command(
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.CopySheet(
                        new SheetName("Budget Copy"), new SheetCopyPosition.AtIndex(1)))));
    WorkbookSheetCommand.SetActiveSheet setActiveSheet =
        cast(
            WorkbookSheetCommand.SetActiveSheet.class,
            command(
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget Copy")),
                    new WorkbookMutationAction.SetActiveSheet())));
    WorkbookSheetCommand.SetSelectedSheets setSelectedSheets =
        cast(
            WorkbookSheetCommand.SetSelectedSheets.class,
            command(
                mutate(
                    new SheetSelector.ByNames(
                        List.of(new SheetName("Budget"), new SheetName("Budget Copy"))),
                    new WorkbookMutationAction.SetSelectedSheets())));
    WorkbookSheetCommand.SetSheetVisibility setSheetVisibility =
        cast(
            WorkbookSheetCommand.SetSheetVisibility.class,
            command(
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.SetSheetVisibility(ExcelSheetVisibility.HIDDEN))));
    WorkbookSheetCommand.SetSheetProtection setSheetProtection =
        cast(
            WorkbookSheetCommand.SetSheetProtection.class,
            command(
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.SetSheetProtection(protectionSettings()))));
    WorkbookSheetCommand.ClearSheetProtection clearSheetProtection =
        cast(
            WorkbookSheetCommand.ClearSheetProtection.class,
            command(
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.ClearSheetProtection())));

    ExcelSheetCopyPosition.AtIndex position =
        assertInstanceOf(ExcelSheetCopyPosition.AtIndex.class, copySheet.position());

    assertEquals("Budget", copySheet.sourceSheetName());
    assertEquals("Budget Copy", copySheet.newSheetName());
    assertEquals(1, position.targetIndex());
    assertEquals("Budget Copy", setActiveSheet.sheetName());
    assertEquals(List.of("Budget", "Budget Copy"), setSelectedSheets.sheetNames());
    assertEquals(ExcelSheetVisibility.HIDDEN, setSheetVisibility.visibility());
    assertEquals(excelProtectionSettings(), setSheetProtection.protection());
    assertEquals("Budget", clearSheetProtection.sheetName());
  }

  private static ConditionalFormattingSpec commandTranslationConditionalFormattingBlock() {
    return new ConditionalFormattingSpec(
        List.of(
            new ConditionalFormattingRuleSpec.FormulaRule(
                "A1>0",
                false,
                java.util.Optional.of(
                    new DifferentialStylePatch(
                        Optional.empty(),
                        Optional.of(true),
                        Optional.empty(),
                        Optional.empty(),
                        java.util.Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        java.util.Optional.empty(),
                        java.util.Optional.empty())))));
  }
}
