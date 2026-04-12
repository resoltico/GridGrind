package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.excel.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.ExcelBinaryData;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.ExcelFormulaCellTarget;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import dev.erst.gridgrind.excel.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.protocol.dto.ChartInput;
import dev.erst.gridgrind.protocol.dto.CommentInput;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.protocol.dto.DataValidationInput;
import dev.erst.gridgrind.protocol.dto.DataValidationRuleInput;
import dev.erst.gridgrind.protocol.dto.DifferentialStyleInput;
import dev.erst.gridgrind.protocol.dto.DrawingAnchorInput;
import dev.erst.gridgrind.protocol.dto.DrawingMarkerInput;
import dev.erst.gridgrind.protocol.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.protocol.dto.FormulaCellTargetInput;
import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.dto.HyperlinkTarget;
import dev.erst.gridgrind.protocol.dto.NamedRangeScope;
import dev.erst.gridgrind.protocol.dto.NamedRangeSelection;
import dev.erst.gridgrind.protocol.dto.NamedRangeTarget;
import dev.erst.gridgrind.protocol.dto.PictureDataInput;
import dev.erst.gridgrind.protocol.dto.PictureInput;
import dev.erst.gridgrind.protocol.dto.RangeSelection;
import dev.erst.gridgrind.protocol.dto.ShapeInput;
import dev.erst.gridgrind.protocol.dto.SheetCopyPosition;
import dev.erst.gridgrind.protocol.dto.SheetProtectionSettings;
import dev.erst.gridgrind.protocol.dto.SheetSelection;
import dev.erst.gridgrind.protocol.dto.TableInput;
import dev.erst.gridgrind.protocol.dto.TableSelection;
import dev.erst.gridgrind.protocol.dto.TableStyleInput;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for SequenceIntrospection operation and command labeling. */
class SequenceIntrospectionTest {
  private static final String PNG_PIXEL_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=";

  @Test
  void reportsWaveThreeOperationKinds() {
    assertEquals(
        "COPY_SHEET",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.CopySheet(
                "Budget", "Budget Copy", new SheetCopyPosition.AppendAtEnd())));
    assertEquals(
        "SET_ACTIVE_SHEET",
        SequenceIntrospection.operationKind(new WorkbookOperation.SetActiveSheet("Budget")));
    assertEquals(
        "SET_SELECTED_SHEETS",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetSelectedSheets(List.of("Budget", "Archive"))));
    assertEquals(
        "SET_SHEET_VISIBILITY",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetSheetVisibility("Budget", ExcelSheetVisibility.HIDDEN)));
    assertEquals(
        "SET_SHEET_PROTECTION",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetSheetProtection("Budget", protocolProtectionSettings())));
    assertEquals(
        "CLEAR_SHEET_PROTECTION",
        SequenceIntrospection.operationKind(new WorkbookOperation.ClearSheetProtection("Budget")));
    assertEquals(
        "SET_HYPERLINK",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetHyperlink(
                "Budget", "A1", new HyperlinkTarget.Url("https://example.com"))));
    assertEquals(
        "CLEAR_HYPERLINK",
        SequenceIntrospection.operationKind(new WorkbookOperation.ClearHyperlink("Budget", "A1")));
    assertEquals(
        "SET_COMMENT",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetComment(
                "Budget", "A1", new CommentInput("Review", "GridGrind", false))));
    assertEquals(
        "CLEAR_COMMENT",
        SequenceIntrospection.operationKind(new WorkbookOperation.ClearComment("Budget", "A1")));
    assertEquals(
        "SET_PICTURE",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetPicture("Budget", protocolPictureInput())));
    assertEquals(
        "SET_SHAPE",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetShape("Budget", protocolShapeInput())));
    assertEquals(
        "SET_EMBEDDED_OBJECT",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetEmbeddedObject("Budget", protocolEmbeddedObjectInput())));
    assertEquals(
        "SET_CHART",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetChart("Budget", protocolChartInput())));
    assertEquals(
        "SET_DRAWING_OBJECT_ANCHOR",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetDrawingObjectAnchor(
                "Budget", "OpsPicture", protocolAnchor())));
    assertEquals(
        "DELETE_DRAWING_OBJECT",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.DeleteDrawingObject("Budget", "OpsPicture")));
    assertEquals(
        "SET_NAMED_RANGE",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetNamedRange(
                "BudgetTotal",
                new NamedRangeScope.Workbook(),
                new NamedRangeTarget("Budget", "B4"))));
    assertEquals(
        "DELETE_NAMED_RANGE",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.DeleteNamedRange("BudgetTotal", new NamedRangeScope.Workbook())));
    assertEquals(
        "SET_DATA_VALIDATION",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetDataValidation(
                "Budget",
                "A2:A5",
                new DataValidationInput(
                    new DataValidationRuleInput.WholeNumber(
                        ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
                    false,
                    false,
                    null,
                    null))));
    assertEquals(
        "CLEAR_DATA_VALIDATIONS",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.ClearDataValidations(
                "Budget", new RangeSelection.Selected(List.of("A2:A5")))));
    assertEquals(
        "SET_CONDITIONAL_FORMATTING",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetConditionalFormatting(
                "Budget",
                new ConditionalFormattingBlockInput(
                    List.of("A2:A5"),
                    List.of(
                        new ConditionalFormattingRuleInput.FormulaRule(
                            "A2>0",
                            true,
                            new DifferentialStyleInput(
                                "0.00", true, null, null, null, null, null, null, null)))))));
    assertEquals(
        "CLEAR_CONDITIONAL_FORMATTING",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.ClearConditionalFormatting(
                "Budget", new RangeSelection.Selected(List.of("A2:A5")))));
    assertEquals(
        "SET_AUTOFILTER",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetAutofilter("Budget", "E1:F4")));
    assertEquals(
        "CLEAR_AUTOFILTER",
        SequenceIntrospection.operationKind(new WorkbookOperation.ClearAutofilter("Budget")));
    assertEquals(
        "SET_TABLE",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.SetTable(
                new TableInput(
                    "BudgetTable",
                    "Budget",
                    "A1:C4",
                    false,
                    new TableStyleInput.Named("TableStyleMedium2", false, false, true, false)))));
    assertEquals(
        "DELETE_TABLE",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.DeleteTable("BudgetTable", "Budget")));
    assertEquals(
        "EVALUATE_FORMULA_CELLS",
        SequenceIntrospection.operationKind(
            new WorkbookOperation.EvaluateFormulaCells(
                List.of(new FormulaCellTargetInput("Budget", "C2")))));
    assertEquals(
        "CLEAR_FORMULA_CACHES",
        SequenceIntrospection.operationKind(new WorkbookOperation.ClearFormulaCaches()));

    assertEquals(
        1L,
        SequenceIntrospection.operationKinds(
                List.of(
                    new WorkbookOperation.SetHyperlink(
                        "Budget", "A1", new HyperlinkTarget.Url("https://example.com"))))
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
        "EVALUATE_FORMULA_CELLS",
        SequenceIntrospection.commandKind(
            new WorkbookCommand.EvaluateFormulaCells(
                List.of(new ExcelFormulaCellTarget("Budget", "C2")))));
    assertEquals(
        "CLEAR_FORMULA_CACHES",
        SequenceIntrospection.commandKind(new WorkbookCommand.ClearFormulaCaches()));

    assertEquals(
        1L,
        SequenceIntrospection.commandKinds(
                List.of(
                    new WorkbookCommand.SetComment(
                        "Budget", "A1", new ExcelComment("Review", "GridGrind", false))))
            .get("SET_COMMENT"));
  }

  @Test
  void reportsReadKindsAndReadCount() {
    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.None(),
            List.of(),
            List.of(
                new WorkbookReadOperation.GetWorkbookSummary("summary"),
                new WorkbookReadOperation.GetWorkbookProtection("workbook-protection"),
                new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")),
                new WorkbookReadOperation.GetDrawingObjects("drawing-objects", "Budget"),
                new WorkbookReadOperation.GetDrawingObjectPayload(
                    "drawing-payload", "Budget", "OpsPicture"),
                new WorkbookReadOperation.GetCharts("charts", "Budget"),
                new WorkbookReadOperation.GetDataValidations(
                    "validations", "Budget", new RangeSelection.All()),
                new WorkbookReadOperation.GetConditionalFormatting(
                    "conditional-formatting", "Budget", new RangeSelection.All()),
                new WorkbookReadOperation.GetAutofilters("autofilters", "Budget"),
                new WorkbookReadOperation.GetTables("tables", new TableSelection.All()),
                new WorkbookReadOperation.GetFormulaSurface("formulas", new SheetSelection.All()),
                new WorkbookReadOperation.AnalyzeDataValidationHealth(
                    "data-validation-health", new SheetSelection.All()),
                new WorkbookReadOperation.AnalyzeConditionalFormattingHealth(
                    "conditional-formatting-health", new SheetSelection.All()),
                new WorkbookReadOperation.AnalyzeAutofilterHealth(
                    "autofilter-health", new SheetSelection.All()),
                new WorkbookReadOperation.AnalyzeTableHealth(
                    "table-health", new TableSelection.All()),
                new WorkbookReadOperation.AnalyzeNamedRangeHealth(
                    "named-range-health", new NamedRangeSelection.All()),
                new WorkbookReadOperation.AnalyzeWorkbookFindings("workbook-findings")));

    assertEquals(17, SequenceIntrospection.readCount(request));
    assertEquals(1L, SequenceIntrospection.readKinds(request.reads()).get("GET_WORKBOOK_SUMMARY"));
    assertEquals(
        1L, SequenceIntrospection.readKinds(request.reads()).get("GET_WORKBOOK_PROTECTION"));
    assertEquals(1L, SequenceIntrospection.readKinds(request.reads()).get("GET_DRAWING_OBJECTS"));
    assertEquals(
        1L, SequenceIntrospection.readKinds(request.reads()).get("GET_DRAWING_OBJECT_PAYLOAD"));
    assertEquals(1L, SequenceIntrospection.readKinds(request.reads()).get("GET_CHARTS"));
    assertEquals(1L, SequenceIntrospection.readKinds(request.reads()).get("GET_DATA_VALIDATIONS"));
    assertEquals(
        1L, SequenceIntrospection.readKinds(request.reads()).get("GET_CONDITIONAL_FORMATTING"));
    assertEquals(1L, SequenceIntrospection.readKinds(request.reads()).get("GET_AUTOFILTERS"));
    assertEquals(1L, SequenceIntrospection.readKinds(request.reads()).get("GET_TABLES"));
    assertEquals(1L, SequenceIntrospection.readKinds(request.reads()).get("GET_FORMULA_SURFACE"));
    assertEquals(
        1L, SequenceIntrospection.readKinds(request.reads()).get("ANALYZE_DATA_VALIDATION_HEALTH"));
    assertEquals(
        1L,
        SequenceIntrospection.readKinds(request.reads())
            .get("ANALYZE_CONDITIONAL_FORMATTING_HEALTH"));
    assertEquals(
        1L, SequenceIntrospection.readKinds(request.reads()).get("ANALYZE_AUTOFILTER_HEALTH"));
    assertEquals(1L, SequenceIntrospection.readKinds(request.reads()).get("ANALYZE_TABLE_HEALTH"));
    assertEquals(
        1L, SequenceIntrospection.readKinds(request.reads()).get("ANALYZE_NAMED_RANGE_HEALTH"));
    assertEquals(
        1L, SequenceIntrospection.readKinds(request.reads()).get("ANALYZE_WORKBOOK_FINDINGS"));
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
        new PictureDataInput(ExcelPictureFormat.PNG, PNG_PIXEL_BASE64),
        protocolAnchor(),
        "Queue preview");
  }

  private static ShapeInput protocolShapeInput() {
    return new ShapeInput(
        "OpsShape", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, protocolAnchor(), "rect", "Queue");
  }

  private static EmbeddedObjectInput protocolEmbeddedObjectInput() {
    return new EmbeddedObjectInput(
        "OpsEmbed",
        "Ops payload",
        "ops-payload.txt",
        "open",
        "R3JpZEdyaW5kIHBheWxvYWQ=",
        new PictureDataInput(ExcelPictureFormat.PNG, PNG_PIXEL_BASE64),
        protocolAnchor());
  }

  private static ChartInput.Bar protocolChartInput() {
    return new ChartInput.Bar(
        "OpsChart",
        protocolAnchor(),
        new ChartInput.Title.Text("Roadmap"),
        new ChartInput.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        true,
        ExcelChartBarDirection.COLUMN,
        List.of(
            new ChartInput.Series(
                new ChartInput.Title.Text("Actual"),
                new ChartInput.DataSource("Budget!$A$2:$A$4"),
                new ChartInput.DataSource("Budget!$B$2:$B$4"))));
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

  private static ExcelChartDefinition.Bar excelChartDefinition() {
    return new ExcelChartDefinition.Bar(
        "OpsChart",
        excelAnchor(),
        new ExcelChartDefinition.Title.Text("Roadmap"),
        new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        true,
        ExcelChartBarDirection.COLUMN,
        List.of(
            new ExcelChartDefinition.Series(
                new ExcelChartDefinition.Title.Text("Actual"),
                new ExcelChartDefinition.DataSource("Budget!$A$2:$A$4"),
                new ExcelChartDefinition.DataSource("Budget!$B$2:$B$4"))));
  }
}
