package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.protocol.CommentInput;
import dev.erst.gridgrind.protocol.DataValidationInput;
import dev.erst.gridgrind.protocol.DataValidationRuleInput;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.HyperlinkTarget;
import dev.erst.gridgrind.protocol.NamedRangeSelection;
import dev.erst.gridgrind.protocol.NamedRangeScope;
import dev.erst.gridgrind.protocol.NamedRangeTarget;
import dev.erst.gridgrind.protocol.RangeSelection;
import dev.erst.gridgrind.protocol.SheetSelection;
import dev.erst.gridgrind.protocol.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.WorkbookOperation;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for SequenceIntrospection operation and command labeling. */
class SequenceIntrospectionTest {
  @Test
  void reportsWaveThreeOperationKinds() {
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
            new WorkbookOperation.DeleteNamedRange(
                "BudgetTotal", new NamedRangeScope.Workbook())));
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
                new WorkbookReadOperation.GetFormulaSurface(
                    "formulas", new SheetSelection.All()),
                new WorkbookReadOperation.AnalyzeDataValidationHealth(
                    "data-validation-health", new SheetSelection.All()),
                new WorkbookReadOperation.AnalyzeNamedRangeHealth(
                    "named-range-health", new NamedRangeSelection.All()),
                new WorkbookReadOperation.AnalyzeWorkbookFindings("workbook-findings")));

    assertEquals(7, SequenceIntrospection.readCount(request));
    assertEquals(
        1L,
        SequenceIntrospection.readKinds(request.reads()).get("GET_WORKBOOK_SUMMARY"));
    assertEquals(
        1L,
        SequenceIntrospection.readKinds(request.reads()).get("GET_DATA_VALIDATIONS"));
    assertEquals(
        1L,
        SequenceIntrospection.readKinds(request.reads()).get("GET_FORMULA_SURFACE"));
    assertEquals(
        1L,
        SequenceIntrospection.readKinds(request.reads()).get("ANALYZE_DATA_VALIDATION_HEALTH"));
    assertEquals(
        1L,
        SequenceIntrospection.readKinds(request.reads()).get("ANALYZE_NAMED_RANGE_HEALTH"));
    assertEquals(
        1L,
        SequenceIntrospection.readKinds(request.reads()).get("ANALYZE_WORKBOOK_FINDINGS"));
  }
}
