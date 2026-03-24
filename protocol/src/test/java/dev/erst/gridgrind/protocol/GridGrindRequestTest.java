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

    GridGrindRequest.WorkbookAnalysisRequest analysisRequest =
        new GridGrindRequest.WorkbookAnalysisRequest(null);
    assertEquals(List.of(), analysisRequest.sheets());
  }

  @Test
  void copiesAndValidatesNestedCollections() {
    List<WorkbookOperation> operations =
        new ArrayList<>(List.of(new WorkbookOperation.EnsureSheet("Budget")));
    List<GridGrindRequest.SheetInspectionRequest> sheets =
        new ArrayList<>(
            List.of(
                new GridGrindRequest.SheetInspectionRequest("Budget", List.of("A1"), null, null)));

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.None(),
            operations,
            new GridGrindRequest.WorkbookAnalysisRequest(sheets));

    operations.clear();
    sheets.clear();

    assertEquals(1, request.operations().size());
    assertEquals(1, request.analysis().sheets().size());

    List<WorkbookOperation> nullOperationList = new ArrayList<>();
    nullOperationList.add(null);
    List<GridGrindRequest.SheetInspectionRequest> nullSheetList = new ArrayList<>();
    nullSheetList.add(null);

    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindRequest(
                new GridGrindRequest.WorkbookSource.New(), null, nullOperationList, null));
    assertThrows(
        NullPointerException.class,
        () -> new GridGrindRequest.WorkbookAnalysisRequest(nullSheetList));
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
        NullPointerException.class, () -> new GridGrindRequest.WorkbookPersistence.SaveAs(null));
    assertThrows(
        IllegalArgumentException.class, () -> new GridGrindRequest.WorkbookPersistence.SaveAs(""));

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
  }
}
