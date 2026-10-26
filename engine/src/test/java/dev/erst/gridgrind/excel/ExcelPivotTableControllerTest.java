package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableController;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableSelection;
import java.io.IOException;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.junit.jupiter.api.Test;

/** Tests for direct pivot-table authoring, introspection, and health analysis. */
class ExcelPivotTableControllerTest {
  private final ExcelPivotTableController controller = new ExcelPivotTableController();

  @Test
  void setPivotTable_roundTripsRangeBackedPivotAcrossSaveAndLoad() throws IOException {
    var workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-pivot-range-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("Report");

      controller.setPivotTable(
          workbook,
          definition(
              "Sales Pivot 2026",
              "Report",
              new PivotTableSpec.PivotSourceSpec.Range(
                  new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                  new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
              "C5",
              List.of(),
              List.of("Region"),
              List.of("Stage")));

      PivotTableReport.Supported snapshot =
          assertInstanceOf(
              PivotTableReport.Supported.class,
              controller
                  .pivotTables(
                      workbook, new ExcelPivotTableSelection.ByNames(List.of("Sales Pivot 2026")))
                  .getFirst());

      assertEquals("Sales Pivot 2026", snapshot.name());
      assertEquals("Report", snapshot.sheetName().value());
      assertEquals("C5", snapshot.anchor().topLeftAddress().value());
      PivotTableReport.PivotSourceReport.Range source =
          assertInstanceOf(PivotTableReport.PivotSourceReport.Range.class, snapshot.source());
      assertEquals("Data", source.sheetName().value());
      assertEquals("A1:D5", source.range().value());
      assertEquals(List.of("Region"), fieldNames(snapshot.rowLabels()));
      assertEquals(List.of("Stage"), fieldNames(snapshot.columnLabels()));
      assertEquals(List.of(), fieldNames(snapshot.reportFilters()));
      assertEquals("Amount", snapshot.dataFields().getFirst().sourceColumnName());
      assertEquals(
          ExcelPivotDataConsolidateFunction.SUM, snapshot.dataFields().getFirst().function());
      assertEquals("Total Amount", snapshot.dataFields().getFirst().displayName());
      assertEquals(Optional.of("#,##0.00"), snapshot.dataFields().getFirst().valueFormat());
      assertTrue(
          controller
              .pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All())
              .isEmpty());

      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      PivotTableReport.Supported snapshot =
          assertInstanceOf(
              PivotTableReport.Supported.class,
              controller.pivotTables(reopened, new ExcelPivotTableSelection.All()).getFirst());

      assertEquals("Sales Pivot 2026", snapshot.name());
      assertEquals("C5", snapshot.anchor().topLeftAddress().value());
      assertTrue(
          controller
              .pivotTableHealthFindings(reopened, new ExcelPivotTableSelection.All())
              .isEmpty());
    }
  }

  @Test
  void setPivotTable_supportsNamedRangeTableSourcesAndReportFilters() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("NamedReport");
      workbook.getOrCreateSheet("TableReport");
      workbook
          .names()
          .setNamedRange(
              new ExcelNamedRangeDefinition(
                  "PivotSource",
                  new ExcelNamedRangeScope.WorkbookScope(),
                  ExcelNamedRangeTarget.range("Data", "A1:D5")));
      workbook
          .tables()
          .setTable(
              new ResolvedTable(
                  new TableName("SalesTable"),
                  new SheetName("Data"),
                  new RangeRef("A1:D5"),
                  false,
                  new TableStyleSpec.None()));

      controller.setPivotTable(
          workbook,
          definition(
              "Named Pivot",
              "NamedReport",
              new PivotTableSpec.PivotSourceSpec.NamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("PivotSource")),
              "A3",
              List.of("Owner"),
              List.of("Region"),
              List.of()));
      controller.setPivotTable(
          workbook,
          definition(
              "Table Pivot",
              "TableReport",
              new PivotTableSpec.PivotSourceSpec.Table(
                  new dev.erst.gridgrind.core.vocabulary.TableName("SalesTable")),
              "F3",
              List.of(),
              List.of("Stage"),
              List.of()));

      List<PivotTableReport> snapshots =
          controller.pivotTables(
              workbook,
              new ExcelPivotTableSelection.ByNames(List.of("Named Pivot", "Table Pivot")));

      PivotTableReport.Supported namedPivot =
          assertInstanceOf(PivotTableReport.Supported.class, snapshots.getFirst());
      PivotTableReport.Supported tablePivot =
          assertInstanceOf(PivotTableReport.Supported.class, snapshots.get(1));

      assertInstanceOf(PivotTableReport.PivotSourceReport.NamedRange.class, namedPivot.source());
      assertEquals(List.of("Owner"), fieldNames(namedPivot.reportFilters()));
      assertEquals("A3", namedPivot.anchor().topLeftAddress().value());
      assertInstanceOf(PivotTableReport.PivotSourceReport.Table.class, tablePivot.source());
      assertEquals(List.of("Stage"), fieldNames(tablePivot.rowLabels()));
      assertTrue(
          controller
              .pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All())
              .isEmpty());
    }
  }

  @Test
  void deletePivotTable_rejectsWrongSheetAndNormalizesCacheIdsAfterDelete() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("ReportA");
      workbook.getOrCreateSheet("ReportB");
      workbook.getOrCreateSheet("ReportC");

      controller.setPivotTable(
          workbook,
          definition(
              "Alpha Pivot",
              "ReportA",
              new PivotTableSpec.PivotSourceSpec.Range(
                  new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                  new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
              "A3",
              List.of(),
              List.of("Region"),
              List.of()));
      controller.setPivotTable(
          workbook,
          definition(
              "Beta Pivot",
              "ReportB",
              new PivotTableSpec.PivotSourceSpec.Range(
                  new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                  new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
              "A3",
              List.of(),
              List.of("Stage"),
              List.of()));

      assertThrows(
          IllegalArgumentException.class,
          () -> controller.deletePivotTable(workbook, "Alpha Pivot", "WrongSheet"));

      controller.deletePivotTable(workbook, "Alpha Pivot", "ReportA");
      controller.setPivotTable(
          workbook,
          definition(
              "Gamma Pivot",
              "ReportC",
              new PivotTableSpec.PivotSourceSpec.Range(
                  new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                  new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
              "A3",
              List.of(),
              List.of("Owner"),
              List.of()));

      List<Long> cacheIds =
          List.of(workbook.xssfWorkbook().getCTWorkbook().getPivotCaches().getPivotCacheArray())
              .stream()
              .map(cache -> cache.getCacheId())
              .toList();
      Set<Long> unique = new LinkedHashSet<>(cacheIds);

      assertEquals(cacheIds.size(), unique.size());
      assertEquals(
          List.of("Beta Pivot", "Gamma Pivot"),
          controller.pivotTables(workbook, new ExcelPivotTableSelection.All()).stream()
              .map(PivotTableReport::name)
              .toList());
    }
  }

  @Test
  void pivotTableHealth_reportsMissingNamesAndBrokenSourcesTruthfully() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("Report");

      controller.setPivotTable(
          workbook,
          definition(
              "Source Pivot",
              "Report",
              new PivotTableSpec.PivotSourceSpec.Range(
                  new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                  new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
              "A3",
              List.of(),
              List.of("Region"),
              List.of()));

      XSSFPivotTable pivotTable = workbook.xssfWorkbook().getPivotTables().getFirst();
      pivotTable.getCTPivotTableDefinition().setName(null);

      List<AnalysisFindingCode> missingNameCodes =
          controller.pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();
      PivotTableReport snapshot =
          controller.pivotTables(workbook, new ExcelPivotTableSelection.All()).getFirst();

      assertTrue(missingNameCodes.contains(AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME));
      assertTrue(snapshot.name().startsWith("_GG_PIVOT_"));

      XSSFPivotCacheDefinition cacheDefinition = cacheDefinition(pivotTable);
      var worksheetSource =
          cacheDefinition.getCTPivotCacheDefinition().getCacheSource().getWorksheetSource();
      worksheetSource.unsetRef();
      worksheetSource.setName("MissingSource");

      List<AnalysisFindingCode> brokenSourceCodes =
          controller.pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();

      assertTrue(brokenSourceCodes.contains(AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME));
      assertTrue(brokenSourceCodes.contains(AnalysisFindingCode.PIVOT_TABLE_BROKEN_SOURCE));
      assertTrue(brokenSourceCodes.contains(AnalysisFindingCode.PIVOT_TABLE_UNSUPPORTED_DETAIL));
    }
  }

  @Test
  void setPivotTable_rejectsReportFilterAnchorsAboveThirdRow() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("Report");

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Filter Pivot",
                          "Report",
                          new PivotTableSpec.PivotSourceSpec.Range(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                              new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
                          "A1",
                          List.of("Owner"),
                          List.of("Region"),
                          List.of())));

      assertTrue(failure.getMessage().contains("row 3 or lower"));
    }
  }

  private PivotTableSpec definition(
      String name,
      String sheetName,
      PivotTableSpec.PivotSourceSpec source,
      String anchor,
      List<String> reportFilters,
      List<String> rowLabels,
      List<String> columnLabels) {
    return new PivotTableSpec(
        name,
        new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName),
        source,
        new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress(anchor)),
        rowLabels,
        columnLabels,
        reportFilters,
        List.of(
            new PivotTableSpec.DataField(
                "Amount",
                ExcelPivotDataConsolidateFunction.SUM,
                "Total Amount",
                Optional.of("#,##0.00"))));
  }

  private void populatePivotSource(ExcelWorkbook workbook, String sheetName) {
    ExcelSheet sheet = workbook.getOrCreateSheet(sheetName);
    sheet
        .cells()
        .setRange(
            "A1:D5",
            List.of(
                List.of(
                    new CellValue.Text("Region"),
                    new CellValue.Text("Stage"),
                    new CellValue.Text("Owner"),
                    new CellValue.Text("Amount")),
                List.of(
                    new CellValue.Text("North"),
                    new CellValue.Text("Plan"),
                    new CellValue.Text("Ada"),
                    new CellValue.NumberValue(10)),
                List.of(
                    new CellValue.Text("North"),
                    new CellValue.Text("Do"),
                    new CellValue.Text("Ada"),
                    new CellValue.NumberValue(15)),
                List.of(
                    new CellValue.Text("South"),
                    new CellValue.Text("Plan"),
                    new CellValue.Text("Lin"),
                    new CellValue.NumberValue(7)),
                List.of(
                    new CellValue.Text("South"),
                    new CellValue.Text("Do"),
                    new CellValue.Text("Lin"),
                    new CellValue.NumberValue(12))));
  }

  private List<String> fieldNames(List<PivotTableReport.Field> fields) {
    return fields.stream().map(PivotTableReport.Field::sourceColumnName).toList();
  }

  private XSSFPivotCacheDefinition cacheDefinition(XSSFPivotTable pivotTable) {
    XSSFPivotCacheDefinition cacheDefinition = pivotTable.getPivotCacheDefinition();
    if (cacheDefinition != null) {
      return cacheDefinition;
    }
    throw new IllegalStateException("Pivot table is missing its cache definition relation.");
  }
}
