package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Direct coverage for stable public contract wording fragments. */
class GridGrindContractTextTest {
  @Test
  void formulaAndPathResolutionSummariesStayStable() {
    assertTrue(GridGrindContractText.formulaAuthoringLimitSummary().contains("LAMBDA/LET"));
    assertTrue(GridGrindContractText.loadedFormulaSupportSummary().contains("UNSUPPORTED_FORMULA"));
    assertTrue(
        GridGrindContractText.stdinExecutionRootRequiredMessage().contains("--execution-root"));
    assertTrue(
        GridGrindContractText.cliFlagPathResolutionSummary().contains("current working directory"));
  }

  @Test
  void catalogHelpersAndSharedSummariesStayStable() {
    assertEquals(
        Set.of(WorkbookMutationAction.EnsureSheet.class, CellMutationAction.AppendRow.class),
        GridGrindContractText.streamingWriteMutationActionClasses());
    assertEquals(
        Set.of(
            WorkbookIntrospectionQuery.GetWorkbookSummary.class,
            SheetIntrospectionQuery.GetSheetSummary.class),
        GridGrindContractText.eventReadInspectionQueryClasses());
    assertTrue(GridGrindContractText.sheetLayoutReadSummary().contains("zoomPercent"));
    assertEquals(
        GridGrindContractText.FORMULA_SURFACE_READ_SUMMARY,
        GridGrindContractText.formulaSurfaceReadSummary());
    assertEquals(
        GridGrindContractText.NAMED_RANGE_SURFACE_READ_SUMMARY,
        GridGrindContractText.namedRangeSurfaceReadSummary());
    assertEquals(
        GridGrindContractText.FORMULA_HEALTH_READ_SUMMARY,
        GridGrindContractText.formulaHealthReadSummary());
    assertEquals(
        GridGrindContractText.NAMED_RANGE_HEALTH_READ_SUMMARY,
        GridGrindContractText.namedRangeHealthReadSummary());
    assertEquals(
        GridGrindContractText.WORKBOOK_FINDINGS_READ_SUMMARY,
        GridGrindContractText.workbookFindingsReadSummary());
    assertTrue(GridGrindContractText.workbookAnalysisFamilyPhrase().contains("formula health"));
    assertTrue(GridGrindContractText.executionModeInputSummary().contains("FULL_XSSF"));
    assertTrue(GridGrindContractText.executionModeInputSummary().contains("EVENT_READ"));
    assertTrue(GridGrindContractText.executionModeInputSummary().contains("STREAMING_WRITE"));
    assertEquals(16L * 1024 * 1024, GridGrindContractText.requestDocumentLimitBytes());
    assertTrue(GridGrindContractText.requestDocumentLimitSummary().contains("16777216 bytes"));
    assertTrue(GridGrindContractText.stepKindSummary().contains("step.type"));
    assertEquals(
        "ENSURE_SHEET",
        GridGrindContractText.mutationActionTypeName(WorkbookMutationAction.EnsureSheet.class));
    assertEquals(
        "GET_WORKBOOK_SUMMARY",
        GridGrindContractText.inspectionQueryTypeName(
            WorkbookIntrospectionQuery.GetWorkbookSummary.class));

    Map<Class<?>, String> inspectionQueryTypeNames =
        GridGrindContractText.typeNamesByClass(InspectionQuery.class);
    assertEquals(
        "GET_WORKBOOK_SUMMARY",
        inspectionQueryTypeNames.get(WorkbookIntrospectionQuery.GetWorkbookSummary.class));
  }

  @Test
  void humanJoinRejectsEmptyListsAndFormatsSinglePairAndSeriesValues() {
    assertEquals("alpha", GridGrindContractText.humanJoin(List.of("alpha")));
    assertEquals("alpha and beta", GridGrindContractText.humanJoin(List.of("alpha", "beta")));
    assertEquals(
        "alpha, beta, and gamma",
        GridGrindContractText.humanJoin(List.of("alpha", "beta", "gamma")));
    assertEquals(
        "alpha and beta", GridGrindContractText.humanJoin(List.of("alpha", "", "beta", " ")));
    assertEquals(
        "values must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> GridGrindContractText.humanJoin(List.of("", " ")))
            .getMessage());
  }
}
