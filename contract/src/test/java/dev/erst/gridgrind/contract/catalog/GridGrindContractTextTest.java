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
    assertTrue(
        GridGrindInspectionContractText.formulaAuthoringLimitSummary().contains("LAMBDA/LET"));
    assertTrue(
        GridGrindInspectionContractText.loadedFormulaSupportSummary()
            .contains("UNSUPPORTED_FORMULA"));
    assertTrue(
        GridGrindRequestSurfaceContractText.stdinExecutionRootRequiredMessage()
            .contains("--execution-root"));
    assertTrue(
        GridGrindRequestSurfaceContractText.cliFlagPathResolutionSummary()
            .contains("current working directory"));
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
    assertTrue(GridGrindInspectionContractText.sheetLayoutReadSummary().contains("zoomPercent"));
    assertEquals(
        GridGrindInspectionContractText.FORMULA_SURFACE_READ_SUMMARY,
        GridGrindInspectionContractText.formulaSurfaceReadSummary());
    assertEquals(
        GridGrindInspectionContractText.NAMED_RANGE_SURFACE_READ_SUMMARY,
        GridGrindInspectionContractText.namedRangeSurfaceReadSummary());
    assertEquals(
        GridGrindInspectionContractText.FORMULA_HEALTH_READ_SUMMARY,
        GridGrindInspectionContractText.formulaHealthReadSummary());
    assertEquals(
        GridGrindInspectionContractText.NAMED_RANGE_HEALTH_READ_SUMMARY,
        GridGrindInspectionContractText.namedRangeHealthReadSummary());
    assertEquals(
        GridGrindInspectionContractText.WORKBOOK_FINDINGS_READ_SUMMARY,
        GridGrindInspectionContractText.workbookFindingsReadSummary());
    assertTrue(
        GridGrindInspectionContractText.workbookAnalysisFamilyPhrase().contains("formula health"));
    assertTrue(GridGrindContractText.executionModeInputSummary().contains("FULL_XSSF"));
    assertTrue(GridGrindContractText.executionModeInputSummary().contains("EVENT_READ"));
    assertTrue(GridGrindContractText.executionModeInputSummary().contains("STREAMING_WRITE"));
    assertEquals(
        16L * 1024 * 1024, GridGrindRequestSurfaceContractText.requestDocumentLimitBytes());
    assertTrue(
        GridGrindRequestSurfaceContractText.requestDocumentLimitSummary()
            .contains("16777216 bytes"));
    assertTrue(GridGrindRequestSurfaceContractText.stepKindSummary().contains("step.type"));
    assertTrue(
        GridGrindRequestSurfaceContractText.encryptedWorkbookTempSecuritySummary()
            .contains("private OS temporary"));
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

  @Test
  void inspectionHumanJoinRejectsEmptyListsAndFormatsSinglePairAndSeriesValues() {
    assertEquals("alpha", GridGrindInspectionContractText.humanJoin(List.of("alpha")));
    assertEquals(
        "alpha and beta", GridGrindInspectionContractText.humanJoin(List.of("alpha", "beta")));
    assertEquals(
        "alpha, beta, and gamma",
        GridGrindInspectionContractText.humanJoin(List.of("alpha", "beta", "gamma")));
    assertEquals(
        "alpha and beta",
        GridGrindInspectionContractText.humanJoin(List.of("alpha", "", "beta", " ")));
    assertEquals(
        "values must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> GridGrindInspectionContractText.humanJoin(List.of("", " ")))
            .getMessage());
  }
}
