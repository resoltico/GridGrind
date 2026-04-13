package dev.erst.gridgrind.protocol.dto;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
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
    assertInstanceOf(GridGrindRequest.WorkbookSource.New.class, request.source());
    assertInstanceOf(GridGrindRequest.WorkbookPersistence.None.class, request.persistence());
    assertNull(request.executionMode());
    assertEquals(List.of(), request.operations());
    assertEquals(List.of(), request.reads());
  }

  @Test
  void copiesAndValidatesNestedCollections() {
    List<WorkbookOperation> operations =
        new ArrayList<>(List.of(new WorkbookOperation.EnsureSheet("Budget")));
    List<WorkbookReadOperation> reads =
        new ArrayList<>(
            List.of(
                new WorkbookReadOperation.GetWorkbookSummary("workbook"),
                new WorkbookReadOperation.GetSheetSummary("sheet-summary", "Budget")));

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.None(),
            operations,
            reads);

    operations.clear();
    reads.clear();

    assertEquals(1, request.operations().size());
    assertEquals(2, request.reads().size());

    List<WorkbookOperation> nullOperationList = new ArrayList<>();
    nullOperationList.add(null);
    List<WorkbookReadOperation> nullReadList = new ArrayList<>();
    nullReadList.add(null);

    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindRequest(
                new GridGrindRequest.WorkbookSource.New(), null, nullOperationList, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindRequest(
                new GridGrindRequest.WorkbookSource.New(), null, null, nullReadList));
  }

  @Test
  void validatesRequiredSourcePersistenceAndReadFields() {
    assertThrows(NullPointerException.class, () -> new GridGrindRequest(null, null, null, null));
    assertEquals(
        GridGrindProtocolVersion.V1,
        new GridGrindRequest(
                null, new GridGrindRequest.WorkbookSource.New(), null, null, null, null)
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
        NullPointerException.class, () -> new WorkbookReadOperation.GetWorkbookSummary(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadOperation.GetWorkbookSummary(" "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadOperation.GetWorkbookProtection(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadOperation.GetWorkbookProtection(" "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadOperation.GetNamedRanges("ranges", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadOperation.GetCells("cells", "Budget", List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadOperation.GetWindow("window", "Budget", "A1", 0, 2));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadOperation.GetWindow("window", "Budget", "A1", 2, 0));
    assertThrows(IllegalArgumentException.class, () -> new NamedRangeSelection.Selected(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new NamedRangeSelection.Selected(
                List.of(
                    new NamedRangeSelector.ByName("Budget"),
                    new NamedRangeSelector.ByName("Budget"))));
    assertThrows(IllegalArgumentException.class, () -> new SheetSelection.Selected(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new SheetSelection.Selected(List.of("Budget", "Budget")));
    assertThrows(IllegalArgumentException.class, () -> new CellSelection.Selected(List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new CellSelection.Selected(List.of("A1", "A1")));
    assertDoesNotThrow(
        () ->
            new WorkbookReadOperation.GetNamedRangeSurface(
                "surface",
                new NamedRangeSelection.Selected(
                    List.of(new NamedRangeSelector.WorkbookScope("BudgetTotal")))));
  }

  @Test
  void normalizesEmptyFormulaEnvironmentToNull() {
    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            null,
            new FormulaEnvironmentInput(List.of(), null, List.of()),
            null,
            null);

    assertNull(request.formulaEnvironment());
  }

  @Test
  void normalizesDefaultExecutionModeToNull() {
    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            null,
            new ExecutionModeInput(null, null),
            null,
            null,
            null);

    assertNull(request.executionMode());
  }

  @Test
  void preservesNonDefaultExecutionMode() {
    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            null,
            new ExecutionModeInput(ExecutionModeInput.ReadMode.EVENT_READ, null),
            null,
            null,
            null);

    assertEquals(ExecutionModeInput.ReadMode.EVENT_READ, request.executionMode().readMode());
    assertEquals(ExecutionModeInput.WriteMode.FULL_XSSF, request.executionMode().writeMode());
  }

  @Test
  void preservesSecuritySettingsOnExistingSourcesAndPersistenceTargets() {
    GridGrindRequest.WorkbookSource.ExistingFile source =
        new GridGrindRequest.WorkbookSource.ExistingFile(
            "budget.xlsx", new OoxmlOpenSecurityInput("source-pass"));
    GridGrindRequest.WorkbookPersistence.SaveAs persistence =
        new GridGrindRequest.WorkbookPersistence.SaveAs(
            "secured.xlsx",
            new OoxmlPersistenceSecurityInput(
                new OoxmlEncryptionInput("persist-pass", null),
                new OoxmlSignatureInput(
                    "tmp/signing-material.p12", "keystore-pass", null, null, null, null)));

    GridGrindRequest request = new GridGrindRequest(source, persistence, List.of(), List.of());

    assertEquals(
        "source-pass",
        ((GridGrindRequest.WorkbookSource.ExistingFile) request.source()).security().password());
    assertEquals(
        "persist-pass",
        ((GridGrindRequest.WorkbookPersistence.SaveAs) request.persistence())
            .security()
            .encryption()
            .password());
    assertEquals(
        "keystore-pass",
        ((GridGrindRequest.WorkbookPersistence.SaveAs) request.persistence())
            .security()
            .signature()
            .keyPassword());
  }

  @Test
  void rejectsDuplicateRequestIds() {
    List<WorkbookReadOperation> duplicateReads =
        List.of(
            new WorkbookReadOperation.GetWorkbookSummary("summary"),
            new WorkbookReadOperation.GetSheetSummary("summary", "Budget"));

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(), null, null, duplicateReads));

    assertTrue(failure.getMessage().contains("summary"));
  }
}
