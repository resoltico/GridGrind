package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.protocol.CommentInput;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.HyperlinkTarget;
import dev.erst.gridgrind.protocol.NamedRangeSelection;
import dev.erst.gridgrind.protocol.NamedRangeScope;
import dev.erst.gridgrind.protocol.NamedRangeTarget;
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
                new WorkbookReadOperation.AnalyzeFormulaSurface(
                    "formulas", new SheetSelection.All()),
                new WorkbookReadOperation.AnalyzeNamedRangeSurface(
                    "ranges", new NamedRangeSelection.All())));

    assertEquals(4, SequenceIntrospection.readCount(request));
    assertEquals(
        1L,
        SequenceIntrospection.readKinds(request.reads()).get("GET_WORKBOOK_SUMMARY"));
    assertEquals(
        1L,
        SequenceIntrospection.readKinds(request.reads()).get("ANALYZE_FORMULA_SURFACE"));
  }
}
