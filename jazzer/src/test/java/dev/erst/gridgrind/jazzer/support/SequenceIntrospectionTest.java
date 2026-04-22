package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.inspect;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.mutate;
import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelArrayFormulaDefinition;
import dev.erst.gridgrind.excel.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.ExcelBinaryData;
import dev.erst.gridgrind.excel.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelCustomXmlImportDefinition;
import dev.erst.gridgrind.excel.ExcelCustomXmlMappingLocator;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.ExcelPivotTableDefinition;
import dev.erst.gridgrind.excel.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelSignatureLineDefinition;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;
import dev.erst.gridgrind.excel.WorkbookCommand;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;

/** Tests for SequenceIntrospection operation and command labeling. */
class SequenceIntrospectionTest {
  private static final String PNG_PIXEL_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=";

  @Test
  void reportsWaveThreeMutationKinds() {
    assertEquals(
        "COPY_SHEET",
        mutationKind(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.CopySheet("Budget Copy", new SheetCopyPosition.AppendAtEnd()))));
    assertEquals(
        "SET_ACTIVE_SHEET",
        mutationKind(
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.SetActiveSheet())));
    assertEquals(
        "SET_SELECTED_SHEETS",
        mutationKind(
            mutate(
                new SheetSelector.ByNames(List.of("Budget", "Archive")),
                new MutationAction.SetSelectedSheets())));
    assertEquals(
        "SET_SHEET_VISIBILITY",
        mutationKind(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetSheetVisibility(ExcelSheetVisibility.HIDDEN))));
    assertEquals(
        "SET_SHEET_PROTECTION",
        mutationKind(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetSheetProtection(protocolProtectionSettings()))));
    assertEquals(
        "CLEAR_SHEET_PROTECTION",
        mutationKind(
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearSheetProtection())));
    assertEquals(
        "SET_HYPERLINK",
        mutationKind(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetHyperlink(new HyperlinkTarget.Url("https://example.com")))));
    assertEquals(
        "CLEAR_HYPERLINK",
        mutationKind(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearHyperlink())));
    assertEquals(
        "SET_COMMENT",
        mutationKind(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetComment(
                    new CommentInput(TextSourceInput.inline("Review"), "GridGrind", false)))));
    assertEquals(
        "CLEAR_COMMENT",
        mutationKind(
            mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearComment())));
    assertEquals(
        "SET_PICTURE",
        mutationKind(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetPicture(protocolPictureInput()))));
    assertEquals(
        "SET_SHAPE",
        mutationKind(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetShape(protocolShapeInput()))));
    assertEquals(
        "SET_EMBEDDED_OBJECT",
        mutationKind(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetEmbeddedObject(protocolEmbeddedObjectInput()))));
    assertEquals(
        "SET_CHART",
        mutationKind(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetChart(protocolChartInput()))));
    assertEquals(
        "SET_PIVOT_TABLE",
        mutationKind(mutate(new MutationAction.SetPivotTable(protocolPivotTableInput()))));
    assertEquals(
        "SET_DRAWING_OBJECT_ANCHOR",
        mutationKind(
            mutate(
                new DrawingObjectSelector.ByName("Budget", "OpsPicture"),
                new MutationAction.SetDrawingObjectAnchor(protocolAnchor()))));
    assertEquals(
        "DELETE_DRAWING_OBJECT",
        mutationKind(
            mutate(
                new DrawingObjectSelector.ByName("Budget", "OpsPicture"),
                new MutationAction.DeleteDrawingObject())));
    assertEquals(
        "SET_NAMED_RANGE",
        mutationKind(
            mutate(
                new MutationAction.SetNamedRange(
                    "BudgetTotal",
                    new NamedRangeScope.Workbook(),
                    new NamedRangeTarget("Budget", "B4")))));
    assertEquals(
        "DELETE_NAMED_RANGE",
        mutationKind(
            mutate(
                new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                new MutationAction.DeleteNamedRange())));
    assertEquals(
        "SET_DATA_VALIDATION",
        mutationKind(
            mutate(
                new RangeSelector.ByRange("Budget", "A2:A5"),
                new MutationAction.SetDataValidation(
                    new DataValidationInput(
                        new DataValidationRuleInput.WholeNumber(
                            ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
                        false,
                        false,
                        null,
                        null)))));
    assertEquals(
        "CLEAR_DATA_VALIDATIONS",
        mutationKind(
            mutate(
                new RangeSelector.ByRanges("Budget", List.of("A2:A5")),
                new MutationAction.ClearDataValidations())));
    assertEquals(
        "SET_CONDITIONAL_FORMATTING",
        mutationKind(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetConditionalFormatting(
                    new ConditionalFormattingBlockInput(
                        List.of("A2:A5"),
                        List.of(
                            new ConditionalFormattingRuleInput.FormulaRule(
                                "A2>0",
                                true,
                                new DifferentialStyleInput(
                                    "0.00", true, null, null, null, null, null, null, null))))))));
    assertEquals(
        "CLEAR_CONDITIONAL_FORMATTING",
        mutationKind(
            mutate(
                new RangeSelector.ByRanges("Budget", List.of("A2:A5")),
                new MutationAction.ClearConditionalFormatting())));
    assertEquals(
        "SET_AUTOFILTER",
        mutationKind(
            mutate(
                new RangeSelector.ByRange("Budget", "E1:F4"), new MutationAction.SetAutofilter())));
    assertEquals(
        "CLEAR_AUTOFILTER",
        mutationKind(
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearAutofilter())));
    assertEquals(
        "SET_TABLE",
        mutationKind(
            mutate(
                new MutationAction.SetTable(
                    new TableInput(
                        "BudgetTable",
                        "Budget",
                        "A1:C4",
                        false,
                        new TableStyleInput.Named(
                            "TableStyleMedium2", false, false, true, false))))));
    assertEquals(
        "DELETE_TABLE",
        mutationKind(
            mutate(
                new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
                new MutationAction.DeleteTable())));
    assertEquals(
        "DELETE_PIVOT_TABLE",
        mutationKind(
            mutate(
                new PivotTableSelector.ByNameOnSheet("OpsPivot", "Budget"),
                new MutationAction.DeletePivotTable())));
    assertEquals(
        "SET_WORKBOOK_PROTECTION",
        mutationKind(
            mutate(
                new WorkbookSelector.Current(),
                new MutationAction.SetWorkbookProtection(
                    new WorkbookProtectionInput(true, false, false, null, null)))));
    assertEquals(
        "CLEAR_WORKBOOK_PROTECTION",
        mutationKind(
            mutate(new WorkbookSelector.Current(), new MutationAction.ClearWorkbookProtection())));
    assertEquals(
        "SET_SHEET_PRESENTATION",
        mutationKind(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetSheetPresentation(SheetPresentationInput.defaults()))));

    assertEquals(
        1L,
        mutationKinds(
                List.of(
                    mutate(
                        new CellSelector.ByAddress("Budget", "A1"),
                        new MutationAction.SetHyperlink(
                            new HyperlinkTarget.Url("https://example.com")))))
            .get("SET_HYPERLINK"));
  }

  @Test
  void reportsWaveThreeCommandKinds() {
    assertEquals(
        "COPY_SHEET",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.CopySheet(
                "Budget", "Budget Copy", new ExcelSheetCopyPosition.AppendAtEnd())));
    assertEquals(
        "SET_ACTIVE_SHEET",
        SequenceIntrospection.commandKind(new WorkbookCommand.SetActiveSheet("Budget")));
    assertEquals(
        "SET_SELECTED_SHEETS",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetSelectedSheets(List.of("Budget", "Archive"))));
    assertEquals(
        "SET_SHEET_VISIBILITY",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetSheetVisibility("Budget", ExcelSheetVisibility.HIDDEN)));
    assertEquals(
        "SET_SHEET_PROTECTION",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetSheetProtection("Budget", excelProtectionSettings())));
    assertEquals(
        "CLEAR_SHEET_PROTECTION",
        SequenceIntrospection.commandKind(new WorkbookCommand.ClearSheetProtection("Budget")));
    assertEquals(
        "SET_HYPERLINK",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetHyperlink(
                "Budget", "A1", new ExcelHyperlink.Url("https://example.com"))));
    assertEquals(
        "CLEAR_HYPERLINK",
        SequenceIntrospection.commandKind(new WorkbookCommand.ClearHyperlink("Budget", "A1")));
    assertEquals(
        "SET_COMMENT",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetComment(
                "Budget", "A1", new ExcelComment("Review", "GridGrind", false))));
    assertEquals(
        "CLEAR_COMMENT",
        SequenceIntrospection.commandKind(new WorkbookCommand.ClearComment("Budget", "A1")));
    assertEquals(
        "SET_PICTURE",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetPicture("Budget", excelPictureDefinition())));
    assertEquals(
        "SET_SHAPE",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetShape("Budget", excelShapeDefinition())));
    assertEquals(
        "SET_EMBEDDED_OBJECT",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetEmbeddedObject("Budget", excelEmbeddedObjectDefinition())));
    assertEquals(
        "SET_CHART",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetChart("Budget", excelChartDefinition())));
    assertEquals(
        "SET_PIVOT_TABLE",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetPivotTable(excelPivotTableDefinition())));
    assertEquals(
        "SET_ARRAY_FORMULA",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetArrayFormula(
                "Budget", "D2:D4", new ExcelArrayFormulaDefinition("B2:B4*C2:C4"))));
    assertEquals(
        "CLEAR_ARRAY_FORMULA",
        SequenceIntrospection.commandKind(new WorkbookCommand.ClearArrayFormula("Budget", "D2")));
    assertEquals(
        "IMPORT_CUSTOM_XML_MAPPING",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.ImportCustomXmlMapping(
                new ExcelCustomXmlImportDefinition(
                    new ExcelCustomXmlMappingLocator(1L, "BudgetMap"),
                    "<Budget><Owner>Ada</Owner></Budget>"))));
    assertEquals(
        "SET_SIGNATURE_LINE",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetSignatureLine("Budget", excelSignatureLineDefinition())));
    assertEquals(
        "SET_DRAWING_OBJECT_ANCHOR",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetDrawingObjectAnchor("Budget", "OpsPicture", excelAnchor())));
    assertEquals(
        "DELETE_DRAWING_OBJECT",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.DeleteDrawingObject("Budget", "OpsPicture")));
    assertEquals(
        "SET_NAMED_RANGE",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetNamedRange(
                new ExcelNamedRangeDefinition(
                    "BudgetTotal",
                    new ExcelNamedRangeScope.WorkbookScope(),
                    new ExcelNamedRangeTarget("Budget", "B4")))));
    assertEquals(
        "DELETE_NAMED_RANGE",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.DeleteNamedRange(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertEquals(
        "SET_DATA_VALIDATION",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetDataValidation(
                "Budget",
                "A2:A5",
                new ExcelDataValidationDefinition(
                    new ExcelDataValidationRule.WholeNumber(
                        ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
                    false,
                    false,
                    null,
                    null))));
    assertEquals(
        "CLEAR_DATA_VALIDATIONS",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.ClearDataValidations(
                "Budget", new dev.erst.gridgrind.excel.ExcelRangeSelection.All())));
    assertEquals(
        "SET_CONDITIONAL_FORMATTING",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetConditionalFormatting(
                "Budget",
                new dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition(
                    List.of("A2:A5"),
                    List.of(
                        new dev.erst.gridgrind.excel.ExcelConditionalFormattingRule.FormulaRule(
                            "A2>0",
                            true,
                            new dev.erst.gridgrind.excel.ExcelDifferentialStyle(
                                "0.00", true, null, null, null, null, null, null, null)))))));
    assertEquals(
        "CLEAR_CONDITIONAL_FORMATTING",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.ClearConditionalFormatting(
                "Budget", new dev.erst.gridgrind.excel.ExcelRangeSelection.All())));
    assertEquals(
        "SET_AUTOFILTER",
        SequenceIntrospection.commandKind(new WorkbookCommand.SetAutofilter("Budget", "E1:F4")));
    assertEquals(
        "CLEAR_AUTOFILTER",
        SequenceIntrospection.commandKind(new WorkbookCommand.ClearAutofilter("Budget")));
    assertEquals(
        "SET_TABLE",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetTable(
                new ExcelTableDefinition(
                    "BudgetTable",
                    "Budget",
                    "A1:C4",
                    false,
                    new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false)))));
    assertEquals(
        "DELETE_TABLE",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.DeleteTable("BudgetTable", "Budget")));
    assertEquals(
        "DELETE_PIVOT_TABLE",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.DeletePivotTable("OpsPivot", "Budget")));
    assertEquals(
        "SET_WORKBOOK_PROTECTION",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetWorkbookProtection(
                new ExcelWorkbookProtectionSettings(true, false, false, null, null))));
    assertEquals(
        "CLEAR_WORKBOOK_PROTECTION",
        SequenceIntrospection.commandKind(new WorkbookCommand.ClearWorkbookProtection()));
    assertEquals(
        "SET_SHEET_PRESENTATION",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.SetSheetPresentation("Budget", ExcelSheetPresentation.defaults())));

    assertEquals(
        1L,
        SequenceIntrospection.commandKinds(
                List.of(
                    new WorkbookCommand.SetComment(
                        "Budget", "A1", new ExcelComment("Review", "GridGrind", false))))
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
                    new InspectionQuery.GetWorkbookSummary()),
                inspect(
                    "workbook-protection",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookProtection()),
                inspect(
                    "custom-xml-mappings",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetCustomXmlMappings()),
                inspect(
                    "custom-xml-export",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.ExportCustomXmlMapping(
                        new dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator(
                            1L, "BudgetMap"),
                        true,
                        "UTF-8")),
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Budget", List.of("A1")),
                    new InspectionQuery.GetCells()),
                inspect(
                    "array-formulas",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetArrayFormulas()),
                inspect(
                    "drawing-objects",
                    new DrawingObjectSelector.AllOnSheet("Budget"),
                    new InspectionQuery.GetDrawingObjects()),
                inspect(
                    "drawing-payload",
                    new DrawingObjectSelector.ByName("Budget", "OpsPicture"),
                    new InspectionQuery.GetDrawingObjectPayload()),
                inspect(
                    "charts",
                    new ChartSelector.AllOnSheet("Budget"),
                    new InspectionQuery.GetCharts()),
                inspect(
                    "pivots",
                    new PivotTableSelector.ByNames(List.of("OpsPivot")),
                    new InspectionQuery.GetPivotTables()),
                inspect(
                    "validations",
                    new RangeSelector.AllOnSheet("Budget"),
                    new InspectionQuery.GetDataValidations()),
                inspect(
                    "conditional-formatting",
                    new RangeSelector.AllOnSheet("Budget"),
                    new InspectionQuery.GetConditionalFormatting()),
                inspect(
                    "autofilters",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetAutofilters()),
                inspect("tables", new TableSelector.All(), new InspectionQuery.GetTables()),
                inspect(
                    "formulas", new SheetSelector.All(), new InspectionQuery.GetFormulaSurface()),
                inspect(
                    "data-validation-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeDataValidationHealth()),
                inspect(
                    "conditional-formatting-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeConditionalFormattingHealth()),
                inspect(
                    "autofilter-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeAutofilterHealth()),
                inspect(
                    "table-health",
                    new TableSelector.All(),
                    new InspectionQuery.AnalyzeTableHealth()),
                inspect(
                    "pivot-table-health",
                    new PivotTableSelector.All(),
                    new InspectionQuery.AnalyzePivotTableHealth()),
                inspect(
                    "named-range-health",
                    new NamedRangeSelector.All(),
                    new InspectionQuery.AnalyzeNamedRangeHealth()),
                inspect(
                    "workbook-findings",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.AnalyzeWorkbookFindings())));

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
                    new InspectionQuery.GetPackageSecurity())));

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
    return SequenceIntrospection.inspectionKinds(request.inspectionSteps());
  }

  private static List<dev.erst.gridgrind.contract.step.MutationStep> materializeMutations(
      List<ProtocolStepSupport.PendingMutation> mutations) {
    return ProtocolStepSupport.steps(mutations, List.of()).stream()
        .map(dev.erst.gridgrind.contract.step.MutationStep.class::cast)
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

  private static DrawingAnchorInput.TwoCell protocolAnchor() {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(0, 0, 0, 0), new DrawingMarkerInput(2, 3, 0, 0), null);
  }

  private static PictureInput protocolPictureInput() {
    return new PictureInput(
        "OpsPicture",
        new PictureDataInput(
            ExcelPictureFormat.PNG, BinarySourceInput.inlineBase64(PNG_PIXEL_BASE64)),
        protocolAnchor(),
        TextSourceInput.inline("Queue preview"));
  }

  private static ShapeInput protocolShapeInput() {
    return new ShapeInput(
        "OpsShape",
        ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
        protocolAnchor(),
        "rect",
        TextSourceInput.inline("Queue"));
  }

  private static EmbeddedObjectInput protocolEmbeddedObjectInput() {
    return new EmbeddedObjectInput(
        "OpsEmbed",
        "Ops payload",
        "ops-payload.txt",
        "open",
        BinarySourceInput.inlineBase64("R3JpZEdyaW5kIHBheWxvYWQ="),
        new PictureDataInput(
            ExcelPictureFormat.PNG, BinarySourceInput.inlineBase64(PNG_PIXEL_BASE64)),
        protocolAnchor());
  }

  private static ChartInput protocolChartInput() {
    return new ChartInput(
        "OpsChart",
        protocolAnchor(),
        new ChartInput.Title.Text(TextSourceInput.inline("Roadmap")),
        new ChartInput.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        List.of(
            new ChartInput.Bar(
                true,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                null,
                null,
                protocolCategoryAxes(),
                List.of(
                    new ChartInput.Series(
                        new ChartInput.Title.Text(TextSourceInput.inline("Actual")),
                        new ChartInput.DataSource.Reference("Budget!$A$2:$A$4"),
                        new ChartInput.DataSource.Reference("Budget!$B$2:$B$4"),
                        null,
                        null,
                        null,
                        null)))));
  }

  private static PivotTableInput protocolPivotTableInput() {
    return new PivotTableInput(
        "OpsPivot",
        "Budget",
        new PivotTableInput.Source.Range("Budget", "A1:C4"),
        new PivotTableInput.Anchor("F4"),
        List.of("Month"),
        List.of(),
        List.of(),
        List.of(
            new PivotTableInput.DataField(
                "Actual", ExcelPivotDataConsolidateFunction.SUM, "Total Actual", null)));
  }

  private static ExcelDrawingAnchor.TwoCell excelAnchor() {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(0, 0, 0, 0), new ExcelDrawingMarker(2, 3, 0, 0), null);
  }

  private static ExcelPictureDefinition excelPictureDefinition() {
    return new ExcelPictureDefinition(
        "OpsPicture",
        new ExcelBinaryData(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
        ExcelPictureFormat.PNG,
        excelAnchor(),
        "Queue preview");
  }

  private static ExcelShapeDefinition excelShapeDefinition() {
    return new ExcelShapeDefinition(
        "OpsShape", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, excelAnchor(), "rect", "Queue");
  }

  private static ExcelEmbeddedObjectDefinition excelEmbeddedObjectDefinition() {
    return new ExcelEmbeddedObjectDefinition(
        "OpsEmbed",
        "Ops payload",
        "ops-payload.txt",
        "open",
        new ExcelBinaryData("GridGrind payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
        ExcelPictureFormat.PNG,
        new ExcelBinaryData(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
        excelAnchor());
  }

  private static ExcelSignatureLineDefinition excelSignatureLineDefinition() {
    return new ExcelSignatureLineDefinition(
        "OpsSignature",
        excelAnchor(),
        false,
        "Review the budget before signing.",
        "Ada Lovelace",
        "Finance",
        "ada@example.com",
        null,
        "invalid",
        ExcelPictureFormat.PNG,
        new ExcelBinaryData(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64)));
  }

  private static ExcelChartDefinition excelChartDefinition() {
    return new ExcelChartDefinition(
        "OpsChart",
        excelAnchor(),
        new ExcelChartDefinition.Title.Text("Roadmap"),
        new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        List.of(
            new ExcelChartDefinition.Bar(
                true,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                null,
                null,
                excelCategoryAxes(),
                List.of(
                    new ExcelChartDefinition.Series(
                        new ExcelChartDefinition.Title.Text("Actual"),
                        new ExcelChartDefinition.DataSource.Reference("Budget!$A$2:$A$4"),
                        new ExcelChartDefinition.DataSource.Reference("Budget!$B$2:$B$4"),
                        null,
                        null,
                        null,
                        null)))));
  }

  private static List<ChartInput.Axis> protocolCategoryAxes() {
    return List.of(
        new ChartInput.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartInput.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ExcelChartDefinition.Axis> excelCategoryAxes() {
    return List.of(
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartDefinition.Axis(
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
                "Actual", ExcelPivotDataConsolidateFunction.SUM, "Total Actual", null)));
  }
}
