package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.protocol.dto.CommentInput;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.protocol.dto.DataValidationInput;
import dev.erst.gridgrind.protocol.dto.DataValidationRuleInput;
import dev.erst.gridgrind.protocol.dto.DifferentialStyleInput;
import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.dto.HyperlinkTarget;
import dev.erst.gridgrind.protocol.dto.NamedRangeScope;
import dev.erst.gridgrind.protocol.dto.NamedRangeSelection;
import dev.erst.gridgrind.protocol.dto.NamedRangeTarget;
import dev.erst.gridgrind.protocol.dto.RangeSelection;
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
                new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")),
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

    assertEquals(13, SequenceIntrospection.readCount(request));
    assertEquals(1L, SequenceIntrospection.readKinds(request.reads()).get("GET_WORKBOOK_SUMMARY"));
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
}
