package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for GridGrindRequest record construction and validation. */
class GridGrindRequestTest {
  @Test
  void appliesDefaultsForOptionalFields() {
    GridGrindRequest request =
        new GridGrindRequest(new GridGrindRequest.WorkbookSource.New(), null, null, null);

    assertEquals(GridGrindProtocolVersion.V1, request.protocolVersion());
    assertTrue(request.source() instanceof GridGrindRequest.WorkbookSource.New);
    assertTrue(request.persistence() instanceof GridGrindRequest.WorkbookPersistence.None);
    assertEquals(List.of(), request.operations());
    assertEquals(List.of(), request.analysis().sheets());
    assertInstanceOf(
        GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.None.class,
        request.analysis().namedRanges());

    GridGrindRequest.WorkbookAnalysisRequest analysisRequest =
        new GridGrindRequest.WorkbookAnalysisRequest(null);
    assertEquals(List.of(), analysisRequest.sheets());
    assertInstanceOf(
        GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.None.class,
        analysisRequest.namedRanges());
  }

  @Test
  void copiesAndValidatesNestedCollections() {
    List<WorkbookOperation> operations =
        new ArrayList<>(List.of(new WorkbookOperation.EnsureSheet("Budget")));
    List<GridGrindRequest.SheetInspectionRequest> sheets =
        new ArrayList<>(
            List.of(
                new GridGrindRequest.SheetInspectionRequest("Budget", List.of("A1"), null, null)));
    List<NamedRangeSelector> selectors =
        new ArrayList<>(List.of(new NamedRangeSelector.ByName("BudgetTotal")));

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.None(),
            operations,
            new GridGrindRequest.WorkbookAnalysisRequest(
                sheets,
                new GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.Selected(
                    selectors)));

    operations.clear();
    sheets.clear();
    selectors.clear();

    assertEquals(1, request.operations().size());
    assertEquals(1, request.analysis().sheets().size());
    GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.Selected namedRangeInspection =
        (GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.Selected)
            request.analysis().namedRanges();
    assertEquals(1, namedRangeInspection.selectors().size());

    List<WorkbookOperation> nullOperationList = new ArrayList<>();
    nullOperationList.add(null);
    List<GridGrindRequest.SheetInspectionRequest> nullSheetList = new ArrayList<>();
    nullSheetList.add(null);
    List<NamedRangeSelector> nullSelectorList = new ArrayList<>();
    nullSelectorList.add(null);

    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindRequest(
                new GridGrindRequest.WorkbookSource.New(), null, nullOperationList, null));
    assertThrows(
        NullPointerException.class,
        () -> new GridGrindRequest.WorkbookAnalysisRequest(nullSheetList));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.Selected(
                nullSelectorList));
  }

  @Test
  void validatesRequiredSourcePersistenceAndAnalysisFields() {
    assertThrows(NullPointerException.class, () -> new GridGrindRequest(null, null, null, null));
    assertEquals(
        GridGrindProtocolVersion.V1,
        new GridGrindRequest(null, new GridGrindRequest.WorkbookSource.New(), null, null, null)
            .protocolVersion());

    assertThrows(
        NullPointerException.class, () -> new GridGrindRequest.WorkbookSource.ExistingFile(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindRequest.WorkbookSource.ExistingFile(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindRequest.WorkbookSource.ExistingFile("budget.xls"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindRequest.WorkbookSource.ExistingFile("budget.xlsm"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindRequest.WorkbookSource.ExistingFile("budget.xlsb"));
    assertEquals(
        "Budget.XLSX", new GridGrindRequest.WorkbookSource.ExistingFile("Budget.XLSX").path());

    assertThrows(
        NullPointerException.class, () -> new GridGrindRequest.WorkbookPersistence.SaveAs(null));
    assertThrows(
        IllegalArgumentException.class, () -> new GridGrindRequest.WorkbookPersistence.SaveAs(""));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindRequest.WorkbookPersistence.SaveAs("report.xls"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindRequest.WorkbookPersistence.SaveAs("report.xlsm"));
    assertEquals(
        "report.XLSX", new GridGrindRequest.WorkbookPersistence.SaveAs("report.XLSX").path());

    assertThrows(
        NullPointerException.class,
        () -> new GridGrindRequest.SheetInspectionRequest(null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindRequest.SheetInspectionRequest("Budget", List.of(" "), null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindRequest.SheetInspectionRequest("Budget", List.of(), 3, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindRequest.SheetInspectionRequest("Budget", List.of(), 0, 2));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindRequest.SheetInspectionRequest("Budget", List.of(), 2, 0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.Selected(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.Selected(null));
    assertDoesNotThrow(
        () ->
            new GridGrindRequest.WorkbookAnalysisRequest(
                List.of(),
                new GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.All()));
  }
}
