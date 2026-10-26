package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.inspect;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.mutate;
import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.model.chart.ChartAxisSpec;
import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.PictureSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedShape;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.model.layout.SheetPresentationSpec;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlImportDefinition;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableDefinition;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationRule;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for SequenceIntrospection operation and command labeling. */
class SequenceIntrospectionTest {
  private static final String PNG_PIXEL_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC";

  @Test
  void reportsWaveThreeMutationKinds() {
    assertEquals(
        "COPY_SHEET",
        mutationKind(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.CopySheet(
                    new SheetName("Budget Copy"), new SheetCopyPosition.AppendAtEnd()))));
    assertEquals(
        "SET_ACTIVE_SHEET",
        mutationKind(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.SetActiveSheet())));
    assertEquals(
        "SET_SELECTED_SHEETS",
        mutationKind(
            mutate(
                new SheetSelector.ByNames(
                    List.of(new SheetName("Budget"), new SheetName("Archive"))),
                new WorkbookMutationAction.SetSelectedSheets())));
    assertEquals(
        "SET_SHEET_VISIBILITY",
        mutationKind(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.SetSheetVisibility(ExcelSheetVisibility.HIDDEN))));
    assertEquals(
        "SET_SHEET_PROTECTION",
        mutationKind(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.SetSheetProtection(protocolProtectionSettings()))));
    assertEquals(
        "CLEAR_SHEET_PROTECTION",
        mutationKind(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.ClearSheetProtection())));
    assertEquals(
        "SET_HYPERLINK",
        mutationKind(
            mutate(
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                new CellMutationAction.SetHyperlink(
                    new HyperlinkTarget.Url("https://example.com")))));
    assertEquals(
        "CLEAR_HYPERLINK",
        mutationKind(
            mutate(
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                new CellMutationAction.ClearHyperlink())));
    assertEquals(
        "SET_COMMENT",
        mutationKind(
            mutate(
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                new CellMutationAction.SetComment(
                    CommentSpec.plain(TextSourceSpec.inline("Review"), "GridGrind", false)))));
    assertEquals(
        "CLEAR_COMMENT",
        mutationKind(
            mutate(
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                new CellMutationAction.ClearComment())));
    assertEquals(
        "SET_PICTURE",
        mutationKind(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new DrawingMutationAction.SetPicture(protocolPictureInput()))));
    assertEquals(
        "SET_SHAPE",
        mutationKind(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new DrawingMutationAction.SetShape(protocolShapeInput()))));
    assertEquals(
        "SET_EMBEDDED_OBJECT",
        mutationKind(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new DrawingMutationAction.SetEmbeddedObject(protocolEmbeddedObjectInput()))));
    assertEquals(
        "SET_CHART",
        mutationKind(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new DrawingMutationAction.SetChart(protocolChartInput()))));
    assertEquals(
        "SET_PIVOT_TABLE",
        mutationKind(
            mutate(new StructuredMutationAction.SetPivotTable(protocolPivotTableInput()))));
    assertEquals(
        "SET_DRAWING_OBJECT_ANCHOR",
        mutationKind(
            mutate(
                new DrawingObjectSelector.ByName(new SheetName("Budget"), "OpsPicture"),
                new DrawingMutationAction.SetDrawingObjectAnchor(protocolAnchor()))));
    assertEquals(
        "DELETE_DRAWING_OBJECT",
        mutationKind(
            mutate(
                new DrawingObjectSelector.ByName(new SheetName("Budget"), "OpsPicture"),
                new DrawingMutationAction.DeleteDrawingObject())));
    assertEquals(
        "SET_NAMED_RANGE",
        mutationKind(
            mutate(
                new StructuredMutationAction.SetNamedRange(
                    new DefinedName("BudgetTotal"),
                    new NamedRangeScope.Workbook(),
                    NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("B4"))))));
    assertEquals(
        "DELETE_NAMED_RANGE",
        mutationKind(
            mutate(
                new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal")),
                new StructuredMutationAction.DeleteNamedRange())));
    assertEquals(
        "SET_DATA_VALIDATION",
        mutationKind(
            mutate(
                new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A2:A5")),
                new StructuredMutationAction.SetDataValidation(
                    new DataValidationSpec(
                        new DataValidationRuleSpec.WholeNumber(
                            ExcelComparisonOperator.GREATER_OR_EQUAL, "1", Optional.empty()),
                        false,
                        false,
                        Optional.empty(),
                        Optional.empty())))));
    assertEquals(
        "CLEAR_DATA_VALIDATIONS",
        mutationKind(
            mutate(
                new RangeSelector.ByRanges(new SheetName("Budget"), List.of(new RangeRef("A2:A5"))),
                new StructuredMutationAction.ClearDataValidations())));
    assertEquals(
        "SET_CONDITIONAL_FORMATTING",
        mutationKind(
            mutate(
                new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A2:A5")),
                new StructuredMutationAction.SetConditionalFormatting(
                    new ConditionalFormattingSpec(
                        List.of(
                            new ConditionalFormattingRuleSpec.FormulaRule(
                                "A2>0",
                                true,
                                Optional.of(
                                    new DifferentialStylePatch(
                                        Optional.of("0.00"),
                                        Optional.of(true),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty())))))))));
    assertEquals(
        "CLEAR_CONDITIONAL_FORMATTING",
        mutationKind(
            mutate(
                new RangeSelector.ByRanges(new SheetName("Budget"), List.of(new RangeRef("A2:A5"))),
                new StructuredMutationAction.ClearConditionalFormatting())));
    assertEquals(
        "SET_AUTOFILTER",
        mutationKind(
            mutate(
                new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("E1:F4")),
                new StructuredMutationAction.SetAutofilter())));
    assertEquals(
        "CLEAR_AUTOFILTER",
        mutationKind(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new StructuredMutationAction.ClearAutofilter())));
    assertEquals(
        "SET_TABLE",
        mutationKind(
            mutate(
                new StructuredMutationAction.SetTable(
                    TableSpec.withDefaultMetadata(
                        new TableName("BudgetTable"),
                        new SheetName("Budget"),
                        new RangeRef("A1:C4"),
                        false,
                        new TableStyleSpec.Named(
                            "TableStyleMedium2", false, false, true, false))))));
    assertEquals(
        "DELETE_TABLE",
        mutationKind(
            mutate(
                new TableSelector.ByNameOnSheet(
                    new TableName("BudgetTable"), new SheetName("Budget")),
                new StructuredMutationAction.DeleteTable())));
    assertEquals(
        "DELETE_PIVOT_TABLE",
        mutationKind(
            mutate(
                new PivotTableSelector.ByNameOnSheet("OpsPivot", new SheetName("Budget")),
                new StructuredMutationAction.DeletePivotTable())));
    assertEquals(
        "SET_WORKBOOK_PROTECTION",
        mutationKind(
            mutate(
                new WorkbookSelector.Current(),
                new WorkbookMutationAction.SetWorkbookProtection(
                    new WorkbookProtectionSpec(
                        true, false, false, Optional.empty(), Optional.empty())))));
    assertEquals(
        "CLEAR_WORKBOOK_PROTECTION",
        mutationKind(
            mutate(
                new WorkbookSelector.Current(),
                new WorkbookMutationAction.ClearWorkbookProtection())));
    assertEquals(
        "SET_SHEET_PRESENTATION",
        mutationKind(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.SetSheetPresentation(
                    SheetPresentationSpec.defaults()))));

    assertEquals(
        1L,
        mutationKinds(
                List.of(
                    mutate(
                        new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                        new CellMutationAction.SetHyperlink(
                            new HyperlinkTarget.Url("https://example.com")))))
            .get("SET_HYPERLINK"));
  }

  @Test
  void reportsWaveThreeCommandKinds() {
    assertEquals(
        "COPY_SHEET",
        SequenceIntrospection.commandKind(
            new WorkbookSheetCommand.CopySheet(
                "Budget", "Budget Copy", new ExcelSheetCopyPosition.AppendAtEnd())));
    assertEquals(
        "SET_ACTIVE_SHEET",
        SequenceIntrospection.commandKind(new WorkbookSheetCommand.SetActiveSheet("Budget")));
    assertEquals(
        "SET_SELECTED_SHEETS",
        SequenceIntrospection.commandKind(
            new WorkbookSheetCommand.SetSelectedSheets(List.of("Budget", "Archive"))));
    assertEquals(
        "SET_SHEET_VISIBILITY",
        SequenceIntrospection.commandKind(
            new WorkbookSheetCommand.SetSheetVisibility("Budget", ExcelSheetVisibility.HIDDEN)));
    assertEquals(
        "SET_SHEET_PROTECTION",
        SequenceIntrospection.commandKind(
            new WorkbookSheetCommand.SetSheetProtection("Budget", excelProtectionSettings())));
    assertEquals(
        "CLEAR_SHEET_PROTECTION",
        SequenceIntrospection.commandKind(new WorkbookSheetCommand.ClearSheetProtection("Budget")));
    assertEquals(
        "SET_HYPERLINK",
        SequenceIntrospection.commandKind(
            new WorkbookAnnotationCommand.SetHyperlink(
                "Budget", "A1", new ExcelHyperlink.Url("https://example.com"))));
    assertEquals(
        "CLEAR_HYPERLINK",
        SequenceIntrospection.commandKind(
            new WorkbookAnnotationCommand.ClearHyperlink("Budget", "A1")));
    assertEquals(
        "SET_COMMENT",
        SequenceIntrospection.commandKind(
            new WorkbookAnnotationCommand.SetComment(
                "Budget", "A1", new ResolvedComment("Review", "GridGrind", false))));
    assertEquals(
        "CLEAR_COMMENT",
        SequenceIntrospection.commandKind(
            new WorkbookAnnotationCommand.ClearComment("Budget", "A1")));
    assertEquals(
        "SET_PICTURE",
        SequenceIntrospection.commandKind(
            new WorkbookDrawingCommand.SetPicture("Budget", excelPictureDefinition())));
    assertEquals(
        "SET_SHAPE",
        SequenceIntrospection.commandKind(
            new WorkbookDrawingCommand.SetShape("Budget", excelShapeDefinition())));
    assertEquals(
        "SET_EMBEDDED_OBJECT",
        SequenceIntrospection.commandKind(
            new WorkbookDrawingCommand.SetEmbeddedObject(
                "Budget", excelEmbeddedObjectDefinition())));
    assertEquals(
        "SET_CHART",
        SequenceIntrospection.commandKind(
            new WorkbookDrawingCommand.SetChart("Budget", excelChartDefinition())));
    assertEquals(
        "SET_PIVOT_TABLE",
        SequenceIntrospection.commandKind(
            new WorkbookTabularCommand.SetPivotTable(excelPivotTableDefinition())));
    assertEquals(
        "SET_ARRAY_FORMULA",
        SequenceIntrospection.commandKind(
            new WorkbookCellCommand.SetArrayFormula(
                "Budget", "D2:D4", new ResolvedArrayFormula("B2:B4*C2:C4"))));
    assertEquals(
        "CLEAR_ARRAY_FORMULA",
        SequenceIntrospection.commandKind(
            new WorkbookCellCommand.ClearArrayFormula("Budget", "D2")));
    assertEquals(
        "IMPORT_CUSTOM_XML_MAPPING",
        SequenceIntrospection.commandKind(
            new WorkbookMetadataCommand.ImportCustomXmlMapping(
                new ExcelCustomXmlImportDefinition(
                    new CustomXmlMappingLocator(
                        java.util.Optional.of(1L), java.util.Optional.of("BudgetMap")),
                    "<Budget><Owner>Ada</Owner></Budget>"))));
    assertEquals(
        "SET_SIGNATURE_LINE",
        SequenceIntrospection.commandKind(
            new WorkbookDrawingCommand.SetSignatureLine("Budget", excelSignatureLineDefinition())));
    assertEquals(
        "SET_DRAWING_OBJECT_ANCHOR",
        SequenceIntrospection.commandKind(
            new WorkbookDrawingCommand.SetDrawingObjectAnchor(
                "Budget", "OpsPicture", excelAnchor())));
    assertEquals(
        "DELETE_DRAWING_OBJECT",
        SequenceIntrospection.commandKind(
            new WorkbookDrawingCommand.DeleteDrawingObject("Budget", "OpsPicture")));
    assertEquals(
        "SET_NAMED_RANGE",
        SequenceIntrospection.commandKind(
            new WorkbookMetadataCommand.SetNamedRange(
                new ExcelNamedRangeDefinition(
                    "BudgetTotal",
                    new ExcelNamedRangeScope.WorkbookScope(),
                    ExcelNamedRangeTarget.range("Budget", "B4")))));
    assertEquals(
        "DELETE_NAMED_RANGE",
        SequenceIntrospection.commandKind(
            new WorkbookMetadataCommand.DeleteNamedRange(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertEquals(
        "SET_DATA_VALIDATION",
        SequenceIntrospection.commandKind(
            new WorkbookFormattingCommand.SetDataValidation(
                "Budget",
                "A2:A5",
                new ExcelDataValidationDefinition(
                    new ExcelDataValidationRule.WholeNumber(
                        ExcelComparisonOperator.GREATER_OR_EQUAL, "1", Optional.empty()),
                    false,
                    false,
                    Optional.empty(),
                    Optional.empty()))));
    assertEquals(
        "CLEAR_DATA_VALIDATIONS",
        SequenceIntrospection.commandKind(
            new WorkbookFormattingCommand.ClearDataValidations(
                "Budget", new dev.erst.gridgrind.excel.ExcelRangeSelection.All())));
    assertEquals(
        "SET_CONDITIONAL_FORMATTING",
        SequenceIntrospection.commandKind(
            new WorkbookFormattingCommand.SetConditionalFormatting(
                "Budget",
                new dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition(
                    List.of("A2:A5"),
                    List.of(
                        new dev.erst.gridgrind.excel.ExcelConditionalFormattingRule.FormulaRule(
                            "A2>0",
                            true,
                            Optional.of(
                                new dev.erst.gridgrind.excel.ExcelDifferentialStyle(
                                    Optional.of("0.00"),
                                    Optional.of(true),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty()))))))));
    assertEquals(
        "CLEAR_CONDITIONAL_FORMATTING",
        SequenceIntrospection.commandKind(
            new WorkbookFormattingCommand.ClearConditionalFormatting(
                "Budget", new dev.erst.gridgrind.excel.ExcelRangeSelection.All())));
    assertEquals(
        "SET_AUTOFILTER",
        SequenceIntrospection.commandKind(
            new WorkbookTabularCommand.SetAutofilter("Budget", "E1:F4")));
    assertEquals(
        "CLEAR_AUTOFILTER",
        SequenceIntrospection.commandKind(new WorkbookTabularCommand.ClearAutofilter("Budget")));
    assertEquals(
        "SET_TABLE",
        SequenceIntrospection.commandKind(
            new WorkbookTabularCommand.SetTable(
                new ExcelTableDefinition(
                    "BudgetTable",
                    "Budget",
                    "A1:C4",
                    false,
                    new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false)))));
    assertEquals(
        "DELETE_TABLE",
        SequenceIntrospection.commandKind(
            new WorkbookTabularCommand.DeleteTable("BudgetTable", "Budget")));
    assertEquals(
        "DELETE_PIVOT_TABLE",
        SequenceIntrospection.commandKind(
            new WorkbookTabularCommand.DeletePivotTable("OpsPivot", "Budget")));
    assertEquals(
        "SET_WORKBOOK_PROTECTION",
        SequenceIntrospection.commandKind(
            new WorkbookSheetCommand.SetWorkbookProtection(
                new ExcelWorkbookProtectionSettings(
                    true, false, false, Optional.empty(), Optional.empty()))));
    assertEquals(
        "CLEAR_WORKBOOK_PROTECTION",
        SequenceIntrospection.commandKind(new WorkbookSheetCommand.ClearWorkbookProtection()));
    assertEquals(
        "SET_SHEET_PRESENTATION",
        SequenceIntrospection.commandKind(
            new WorkbookLayoutCommand.SetSheetPresentation(
                "Budget", ExcelSheetPresentation.defaults())));

    assertEquals(
        1L,
        SequenceIntrospection.commandKinds(
                List.of(
                    new WorkbookAnnotationCommand.SetComment(
                        "Budget", "A1", new ResolvedComment("Review", "GridGrind", false))))
            .get("SET_COMMENT"));
  }

  @Test
  void reportsInspectionKindsAndInspectionCount() {
    WorkbookPlan request =
        ProtocolStepSupport.request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(),
            List.of(
                inspect(
                    "summary",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary()),
                inspect(
                    "workbook-protection",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookProtection()),
                inspect(
                    "custom-xml-mappings",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetCustomXmlMappings()),
                inspect(
                    "custom-xml-export",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                        new dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator(
                            java.util.Optional.of(1L), java.util.Optional.of("BudgetMap")),
                        true,
                        "UTF-8")),
                inspect(
                    "cells",
                    new CellSelector.ByAddresses(
                        new SheetName("Budget"), List.of(new CellAddress("A1"))),
                    new SheetIntrospectionQuery.GetCells()),
                inspect(
                    "array-formulas",
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new SheetIntrospectionQuery.GetArrayFormulas()),
                inspect(
                    "drawing-objects",
                    new DrawingObjectSelector.AllOnSheet(new SheetName("Budget")),
                    new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
                inspect(
                    "drawing-payload",
                    new DrawingObjectSelector.ByName(new SheetName("Budget"), "OpsPicture"),
                    new WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload()),
                inspect(
                    "charts",
                    new ChartSelector.AllOnSheet(new SheetName("Budget")),
                    new WorkbookAssetIntrospectionQuery.GetCharts()),
                inspect(
                    "pivots",
                    new PivotTableSelector.ByNames(List.of("OpsPivot")),
                    new WorkbookAssetIntrospectionQuery.GetPivotTables()),
                inspect(
                    "validations",
                    new RangeSelector.AllOnSheet(new SheetName("Budget")),
                    new SheetIntrospectionQuery.GetDataValidations()),
                inspect(
                    "conditional-formatting",
                    new RangeSelector.AllOnSheet(new SheetName("Budget")),
                    new SheetIntrospectionQuery.GetConditionalFormatting()),
                inspect(
                    "autofilters",
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new SheetIntrospectionQuery.GetAutofilters()),
                inspect(
                    "tables",
                    new TableSelector.All(),
                    new WorkbookAssetIntrospectionQuery.GetTables()),
                inspect(
                    "formulas",
                    new SheetSelector.All(),
                    new InspectionSurfaceQuery.GetFormulaSurface()),
                inspect(
                    "data-validation-health",
                    new SheetSelector.All(),
                    new InspectionAnalysisQuery.AnalyzeDataValidationHealth()),
                inspect(
                    "conditional-formatting-health",
                    new SheetSelector.All(),
                    new InspectionAnalysisQuery.AnalyzeConditionalFormattingHealth()),
                inspect(
                    "autofilter-health",
                    new SheetSelector.All(),
                    new InspectionAnalysisQuery.AnalyzeAutofilterHealth()),
                inspect(
                    "table-health",
                    new TableSelector.All(),
                    new InspectionAnalysisQuery.AnalyzeTableHealth()),
                inspect(
                    "pivot-table-health",
                    new PivotTableSelector.All(),
                    new InspectionAnalysisQuery.AnalyzePivotTableHealth()),
                inspect(
                    "named-range-health",
                    new NamedRangeSelector.All(),
                    new InspectionAnalysisQuery.AnalyzeNamedRangeHealth()),
                inspect(
                    "workbook-findings",
                    new WorkbookSelector.Current(),
                    new InspectionAnalysisQuery.AnalyzeWorkbookFindings())));

    assertEquals(22, inspectionCount(request));
    assertEquals(1L, inspectionKinds(request).get("GET_WORKBOOK_SUMMARY"));
    assertEquals(1L, inspectionKinds(request).get("GET_WORKBOOK_PROTECTION"));
    assertEquals(1L, inspectionKinds(request).get("GET_CUSTOM_XML_MAPPINGS"));
    assertEquals(1L, inspectionKinds(request).get("EXPORT_CUSTOM_XML_MAPPING"));
    assertEquals(1L, inspectionKinds(request).get("GET_ARRAY_FORMULAS"));
    assertEquals(1L, inspectionKinds(request).get("GET_DRAWING_OBJECTS"));
    assertEquals(1L, inspectionKinds(request).get("GET_DRAWING_OBJECT_PAYLOAD"));
    assertEquals(1L, inspectionKinds(request).get("GET_CHARTS"));
    assertEquals(1L, inspectionKinds(request).get("GET_PIVOT_TABLES"));
    assertEquals(1L, inspectionKinds(request).get("GET_DATA_VALIDATIONS"));
    assertEquals(1L, inspectionKinds(request).get("GET_CONDITIONAL_FORMATTING"));
    assertEquals(1L, inspectionKinds(request).get("GET_AUTOFILTERS"));
    assertEquals(1L, inspectionKinds(request).get("GET_TABLES"));
    assertEquals(1L, inspectionKinds(request).get("GET_FORMULA_SURFACE"));
    assertEquals(1L, inspectionKinds(request).get("ANALYZE_DATA_VALIDATION_HEALTH"));
    assertEquals(1L, inspectionKinds(request).get("ANALYZE_CONDITIONAL_FORMATTING_HEALTH"));
    assertEquals(1L, inspectionKinds(request).get("ANALYZE_AUTOFILTER_HEALTH"));
    assertEquals(1L, inspectionKinds(request).get("ANALYZE_TABLE_HEALTH"));
    assertEquals(1L, inspectionKinds(request).get("ANALYZE_PIVOT_TABLE_HEALTH"));
    assertEquals(1L, inspectionKinds(request).get("ANALYZE_NAMED_RANGE_HEALTH"));
    assertEquals(1L, inspectionKinds(request).get("ANALYZE_WORKBOOK_FINDINGS"));
  }

  @Test
  void inspectionKindsCountPackageSecurityReads() {
    WorkbookPlan request =
        ProtocolStepSupport.request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(),
            List.of(
                inspect(
                    "security",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetPackageSecurity())));

    assertEquals(1, inspectionCount(request));
    assertEquals(1L, inspectionKinds(request).get("GET_PACKAGE_SECURITY"));
  }

  private static String mutationKind(ProtocolStepSupport.PendingMutation mutation) {
    return SequenceIntrospection.mutationKind(ProtocolStepSupport.materializeMutation(mutation, 0));
  }

  private static Map<String, Long> mutationKinds(
      List<ProtocolStepSupport.PendingMutation> mutations) {
    return SequenceIntrospection.mutationKinds(materializeMutations(mutations));
  }

  private static int inspectionCount(WorkbookPlan request) {
    return SequenceIntrospection.inspectionCount(request);
  }

  private static Map<String, Long> inspectionKinds(WorkbookPlan request) {
    return SequenceIntrospection.inspectionKinds(request.stepPartition().inspections());
  }

  private static List<dev.erst.gridgrind.core.plan.MutationStep> materializeMutations(
      List<ProtocolStepSupport.PendingMutation> mutations) {
    return ProtocolStepSupport.steps(mutations, List.of()).stream()
        .map(dev.erst.gridgrind.core.plan.MutationStep.class::cast)
        .toList();
  }

  private static SheetProtectionSettings protocolProtectionSettings() {
    return new SheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }

  private static ExcelSheetProtectionSettings excelProtectionSettings() {
    return new ExcelSheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }

  private static DrawingAnchorSpec.TwoCell protocolAnchor() {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(0, 0, 0, 0),
        new DrawingMarkerSpec(2, 3, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static PictureSpec protocolPictureInput() {
    return new PictureSpec(
        "OpsPicture",
        new PictureDataSpec(
            ExcelPictureFormat.PNG, BinarySourceSpec.inlineBase64(PNG_PIXEL_BASE64)),
        protocolAnchor(),
        Optional.of(TextSourceSpec.inline("Queue preview")));
  }

  private static ShapeSpec protocolShapeInput() {
    return new ShapeSpec.SimpleShape(
        "OpsShape", protocolAnchor(), "rect", Optional.of(TextSourceSpec.inline("Queue")));
  }

  private static EmbeddedObjectSpec protocolEmbeddedObjectInput() {
    return new EmbeddedObjectSpec(
        "OpsEmbed",
        "Ops payload",
        "ops-payload.txt",
        "open",
        BinarySourceSpec.inlineBase64("R3JpZEdyaW5kIHBheWxvYWQ="),
        new PictureDataSpec(
            ExcelPictureFormat.PNG, BinarySourceSpec.inlineBase64(PNG_PIXEL_BASE64)),
        protocolAnchor());
  }

  private static ChartSpec protocolChartInput() {
    return new ChartSpec(
        "OpsChart",
        protocolAnchor(),
        new ChartTitleSpec.Text(TextSourceSpec.inline("Roadmap")),
        new ChartLegendSpec.Visible(ExcelChartLegendPosition.TOP_RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        List.of(
            new ChartPlotSpec.Bar(
                true,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                Optional.empty(),
                Optional.empty(),
                protocolCategoryAxes(),
                List.of(
                    new ChartSeriesSpec(
                        new ChartTitleSpec.Text(TextSourceSpec.inline("Actual")),
                        new ChartDataSourceSpec.Reference("Budget!$A$2:$A$4"),
                        new ChartDataSourceSpec.Reference("Budget!$B$2:$B$4"),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty())))));
  }

  private static PivotTableSpec protocolPivotTableInput() {
    return new PivotTableSpec(
        "OpsPivot",
        new SheetName("Budget"),
        new PivotTableSpec.PivotSourceSpec.Range(new SheetName("Budget"), new RangeRef("A1:C4")),
        new PivotTableSpec.Anchor(new CellAddress("F4")),
        List.of("Month"),
        List.of(),
        List.of(),
        List.of(
            new PivotTableSpec.DataField(
                "Actual",
                ExcelPivotDataConsolidateFunction.SUM,
                "Total Actual",
                Optional.empty())));
  }

  private static DrawingAnchorSpec.TwoCell excelAnchor() {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(0, 0, 0, 0),
        new DrawingMarkerSpec(2, 3, 0, 0),
        dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static ResolvedPicture excelPictureDefinition() {
    return new ResolvedPicture(
        "OpsPicture",
        new BinaryContent(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
        ExcelPictureFormat.PNG,
        excelAnchor(),
        Optional.of("Queue preview"));
  }

  private static ResolvedShape excelShapeDefinition() {
    return new ResolvedShape.SimpleShape("OpsShape", excelAnchor(), "rect", Optional.of("Queue"));
  }

  private static ResolvedEmbeddedObject excelEmbeddedObjectDefinition() {
    return new ResolvedEmbeddedObject(
        "OpsEmbed",
        "Ops payload",
        "ops-payload.txt",
        "open",
        new BinaryContent("GridGrind payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
        ExcelPictureFormat.PNG,
        new BinaryContent(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
        excelAnchor());
  }

  private static ResolvedSignatureLine excelSignatureLineDefinition() {
    return new ResolvedSignatureLine(
        "OpsSignature",
        excelAnchor(),
        false,
        Optional.ofNullable("Review the budget before signing."),
        Optional.ofNullable("Ada Lovelace"),
        Optional.ofNullable("Finance"),
        Optional.ofNullable("ada@example.com"),
        Optional.empty(),
        Optional.ofNullable("invalid"),
        Optional.of(ExcelPictureFormat.PNG),
        Optional.of(new BinaryContent(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64))));
  }

  private static ResolvedChart excelChartDefinition() {
    return new ResolvedChart(
        "OpsChart",
        excelAnchor(),
        new ResolvedChartTitle.Text("Roadmap"),
        new ChartLegendSpec.Visible(ExcelChartLegendPosition.TOP_RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        List.of(
            new ResolvedChart.Bar(
                true,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                Optional.empty(),
                Optional.empty(),
                excelCategoryAxes(),
                List.of(
                    new ResolvedChart.Series(
                        new ResolvedChartTitle.Text("Actual"),
                        new ChartDataSourceSpec.Reference("Budget!$A$2:$A$4"),
                        new ChartDataSourceSpec.Reference("Budget!$B$2:$B$4"),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty())))));
  }

  private static List<ChartAxisSpec> protocolCategoryAxes() {
    return List.of(
        new ChartAxisSpec(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartAxisSpec> excelCategoryAxes() {
    return List.of(
        new ChartAxisSpec(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static ExcelPivotTableDefinition excelPivotTableDefinition() {
    return new ExcelPivotTableDefinition(
        "OpsPivot",
        "Budget",
        new ExcelPivotTableDefinition.Source.Range("Budget", "A1:C4"),
        new ExcelPivotTableDefinition.Anchor("F4"),
        List.of("Month"),
        List.of(),
        List.of(),
        List.of(
            new ExcelPivotTableDefinition.DataField(
                "Actual",
                ExcelPivotDataConsolidateFunction.SUM,
                "Total Actual",
                Optional.empty())));
  }
}
