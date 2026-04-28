package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import java.io.IOException;
import java.util.LinkedHashSet;
import java.util.List;
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

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("Report");

      controller.setPivotTable(
          workbook,
          definition(
              "Sales Pivot 2026",
              "Report",
              new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
              "C5",
              List.of(),
              List.of("Region"),
              List.of("Stage")));

      ExcelPivotTableSnapshot.Supported snapshot =
          assertInstanceOf(
              ExcelPivotTableSnapshot.Supported.class,
              controller
                  .pivotTables(
                      workbook, new ExcelPivotTableSelection.ByNames(List.of("Sales Pivot 2026")))
                  .getFirst());

      assertEquals("Sales Pivot 2026", snapshot.name());
      assertEquals("Report", snapshot.sheetName());
      assertEquals("C5", snapshot.anchor().topLeftAddress());
      ExcelPivotTableSnapshot.Source.Range source =
          assertInstanceOf(ExcelPivotTableSnapshot.Source.Range.class, snapshot.source());
      assertEquals("Data", source.sheetName());
      assertEquals("A1:D5", source.range());
      assertEquals(List.of("Region"), fieldNames(snapshot.rowLabels()));
      assertEquals(List.of("Stage"), fieldNames(snapshot.columnLabels()));
      assertEquals(List.of(), fieldNames(snapshot.reportFilters()));
      assertEquals("Amount", snapshot.dataFields().getFirst().sourceColumnName());
      assertEquals(
          ExcelPivotDataConsolidateFunction.SUM, snapshot.dataFields().getFirst().function());
      assertEquals("Total Amount", snapshot.dataFields().getFirst().displayName());
      assertEquals("#,##0.00", snapshot.dataFields().getFirst().valueFormat());
      assertTrue(
          controller
              .pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All())
              .isEmpty());

      workbook.save(workbookPath);
    }

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      ExcelPivotTableSnapshot.Supported snapshot =
          assertInstanceOf(
              ExcelPivotTableSnapshot.Supported.class,
              controller.pivotTables(reopened, new ExcelPivotTableSelection.All()).getFirst());

      assertEquals("Sales Pivot 2026", snapshot.name());
      assertEquals("C5", snapshot.anchor().topLeftAddress());
      assertTrue(
          controller
              .pivotTableHealthFindings(reopened, new ExcelPivotTableSelection.All())
              .isEmpty());
    }
  }

  @Test
  void setPivotTable_supportsNamedRangeTableSourcesAndReportFilters() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("NamedReport");
      workbook.getOrCreateSheet("TableReport");
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "PivotSource",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Data", "A1:D5")));
      workbook.setTable(
          new ExcelTableDefinition(
              "SalesTable", "Data", "A1:D5", false, new ExcelTableStyle.None()));

      controller.setPivotTable(
          workbook,
          definition(
              "Named Pivot",
              "NamedReport",
              new ExcelPivotTableDefinition.Source.NamedRange("PivotSource"),
              "A3",
              List.of("Owner"),
              List.of("Region"),
              List.of()));
      controller.setPivotTable(
          workbook,
          definition(
              "Table Pivot",
              "TableReport",
              new ExcelPivotTableDefinition.Source.Table("SalesTable"),
              "F3",
              List.of(),
              List.of("Stage"),
              List.of()));

      List<ExcelPivotTableSnapshot> snapshots =
          controller.pivotTables(
              workbook,
              new ExcelPivotTableSelection.ByNames(List.of("Named Pivot", "Table Pivot")));

      ExcelPivotTableSnapshot.Supported namedPivot =
          assertInstanceOf(ExcelPivotTableSnapshot.Supported.class, snapshots.getFirst());
      ExcelPivotTableSnapshot.Supported tablePivot =
          assertInstanceOf(ExcelPivotTableSnapshot.Supported.class, snapshots.get(1));

      assertInstanceOf(ExcelPivotTableSnapshot.Source.NamedRange.class, namedPivot.source());
      assertEquals(List.of("Owner"), fieldNames(namedPivot.reportFilters()));
      assertEquals("A3", namedPivot.anchor().topLeftAddress());
      assertInstanceOf(ExcelPivotTableSnapshot.Source.Table.class, tablePivot.source());
      assertEquals(List.of("Stage"), fieldNames(tablePivot.rowLabels()));
      assertTrue(
          controller
              .pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All())
              .isEmpty());
    }
  }

  @Test
  void deletePivotTable_rejectsWrongSheetAndNormalizesCacheIdsAfterDelete() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("ReportA");
      workbook.getOrCreateSheet("ReportB");
      workbook.getOrCreateSheet("ReportC");

      controller.setPivotTable(
          workbook,
          definition(
              "Alpha Pivot",
              "ReportA",
              new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
              "A3",
              List.of(),
              List.of("Region"),
              List.of()));
      controller.setPivotTable(
          workbook,
          definition(
              "Beta Pivot",
              "ReportB",
              new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
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
              new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
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
              .map(ExcelPivotTableSnapshot::name)
              .toList());
    }
  }

  @Test
  void pivotTableHealth_reportsMissingNamesAndBrokenSourcesTruthfully() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("Report");

      controller.setPivotTable(
          workbook,
          definition(
              "Source Pivot",
              "Report",
              new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
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
      ExcelPivotTableSnapshot snapshot =
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
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
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
                          new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                          "A1",
                          List.of("Owner"),
                          List.of("Region"),
                          List.of())));

      assertTrue(failure.getMessage().contains("row 3 or lower"));
    }
  }

  private ExcelPivotTableDefinition definition(
      String name,
      String sheetName,
      ExcelPivotTableDefinition.Source source,
      String anchor,
      List<String> reportFilters,
      List<String> rowLabels,
      List<String> columnLabels) {
    return new ExcelPivotTableDefinition(
        name,
        sheetName,
        source,
        new ExcelPivotTableDefinition.Anchor(anchor),
        rowLabels,
        columnLabels,
        reportFilters,
        List.of(
            new ExcelPivotTableDefinition.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", "#,##0.00")));
  }

  private void populatePivotSource(ExcelWorkbook workbook, String sheetName) {
    ExcelSheet sheet = workbook.getOrCreateSheet(sheetName);
    sheet.setRange(
        "A1:D5",
        List.of(
            List.of(
                ExcelCellValue.text("Region"),
                ExcelCellValue.text("Stage"),
                ExcelCellValue.text("Owner"),
                ExcelCellValue.text("Amount")),
            List.of(
                ExcelCellValue.text("North"),
                ExcelCellValue.text("Plan"),
                ExcelCellValue.text("Ada"),
                ExcelCellValue.number(10)),
            List.of(
                ExcelCellValue.text("North"),
                ExcelCellValue.text("Do"),
                ExcelCellValue.text("Ada"),
                ExcelCellValue.number(15)),
            List.of(
                ExcelCellValue.text("South"),
                ExcelCellValue.text("Plan"),
                ExcelCellValue.text("Lin"),
                ExcelCellValue.number(7)),
            List.of(
                ExcelCellValue.text("South"),
                ExcelCellValue.text("Do"),
                ExcelCellValue.text("Lin"),
                ExcelCellValue.number(12))));
  }

  private List<String> fieldNames(List<ExcelPivotTableSnapshot.Field> fields) {
    return fields.stream().map(ExcelPivotTableSnapshot.Field::sourceColumnName).toList();
  }

  private XSSFPivotCacheDefinition cacheDefinition(XSSFPivotTable pivotTable) {
    XSSFPivotCacheDefinition cacheDefinition = pivotTable.getPivotCacheDefinition();
    if (cacheDefinition != null) {
      return cacheDefinition;
    }
    throw new IllegalStateException("Pivot table is missing its cache definition relation.");
  }
}
