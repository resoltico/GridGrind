package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.poi.CellBounds;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Verifies Unicode defined names and table names survive authored and existing-workbook reads. */
class ExcelUnicodeDefinedNamePersistenceTest {
  private static final List<UnicodeNameFixture> FIXTURES =
      List.of(
          new UnicodeNameFixture("IeņēmumiRange", "IeņēmumiTable", "A1:B2"),
          new UnicodeNameFixture("ДоходыRange", "ДоходыTable", "D1:E2"),
          new UnicodeNameFixture("収益Range", "収益Table", "G1:H2"));

  @Test
  void authorsAndReopensUnicodeNamedRangesAndTables() throws IOException {
    Path workbookPath =
        ManagedTempFiles.createManagedTempFile("gridgrind-unicode-author-", ".xlsx");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Data");
      populateTableRanges(sheet);
      for (UnicodeNameFixture fixture : FIXTURES) {
        authorFixture(workbook, fixture);
      }
      workbook
          .persistence()
          .save(
              workbookPath,
              WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XssfWorkbook workbook =
        XssfWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals(
          FIXTURES.stream().map(UnicodeNameFixture::namedRangeName).toList(),
          workbook.names().namedRanges().stream()
              .map(namedRange -> namedRange.name().value())
              .toList());
      assertEquals(
          FIXTURES.stream().map(UnicodeNameFixture::tableName).toList(),
          workbook.xssfWorkbook().getSheet("Data").getTables().stream()
              .map(table -> table.getName())
              .toList());
    }
  }

  @Test
  void readsUnicodeNamesAndTablesCreatedByAnExistingWorkbook() throws IOException {
    Path workbookPath =
        ManagedTempFiles.createManagedTempFile("gridgrind-unicode-observed-", ".xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Data");
      populateTableRanges(sheet);
      for (UnicodeNameFixture fixture : FIXTURES) {
        authorObservedFixture(workbook, sheet, fixture);
      }
      try (var output = Files.newOutputStream(workbookPath)) {
        workbook.write(output);
      }
    }

    try (XssfWorkbook workbook =
        XssfWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertTrue(
          workbook.names().namedRanges().stream()
              .map(namedRange -> namedRange.name().value())
              .collect(java.util.stream.Collectors.toSet())
              .containsAll(FIXTURES.stream().map(UnicodeNameFixture::namedRangeName).toList()));
      assertTrue(
          workbook.xssfWorkbook().getSheet("Data").getTables().stream()
              .map(table -> table.getName())
              .toList()
              .containsAll(FIXTURES.stream().map(UnicodeNameFixture::tableName).toList()));
    }
  }

  private static void populateTableRanges(XssfSheet sheet) {
    for (UnicodeNameFixture fixture : FIXTURES) {
      populateTableRange(sheet, fixture.range());
    }
  }

  private static void populateTableRanges(org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
    for (UnicodeNameFixture fixture : FIXTURES) {
      populateObservedTableRange(sheet, fixture.range());
    }
  }

  private static String column(int columnIndex) {
    return org.apache.poi.ss.util.CellReference.convertNumToColString(columnIndex);
  }

  private static void authorFixture(XssfWorkbook workbook, UnicodeNameFixture fixture) {
    workbook
        .names()
        .setNamedRange(
            new StructuredMutationAction.SetNamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName(fixture.namedRangeName()),
                new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef(fixture.range()))));
    new TableAdapter()
        .setTable(
            workbook,
            new ResolvedTable(
                new TableName(fixture.tableName()),
                new SheetName("Data"),
                new RangeRef(fixture.range()),
                false,
                new TableStyleSpec.None()));
  }

  private static void authorObservedFixture(
      XSSFWorkbook workbook,
      org.apache.poi.xssf.usermodel.XSSFSheet sheet,
      UnicodeNameFixture fixture) {
    CellBounds parsed = CellBounds.parse(fixture.range());
    var namedRange = workbook.createName();
    namedRange.setNameName(fixture.namedRangeName());
    namedRange.setRefersToFormula(
        "Data!$" + column(parsed.firstColumn()) + "$1:$" + column(parsed.lastColumn()) + "$2");
    var table = sheet.createTable(new AreaReference(fixture.range(), SpreadsheetVersion.EXCEL2007));
    table.setName(fixture.tableName());
    table.setDisplayName(fixture.tableName());
  }

  private static void populateTableRange(XssfSheet sheet, String range) {
    CellBounds parsed = CellBounds.parse(range);
    sheet
        .cells()
        .setCell(
            new org.apache.poi.ss.util.CellReference(parsed.firstRow(), parsed.firstColumn())
                .formatAsString(),
            new CellValue.Text("Header" + parsed.firstColumn()));
    sheet
        .cells()
        .setCell(
            new org.apache.poi.ss.util.CellReference(parsed.firstRow(), parsed.lastColumn())
                .formatAsString(),
            new CellValue.Text("Value" + parsed.lastColumn()));
    sheet
        .cells()
        .setCell(
            new org.apache.poi.ss.util.CellReference(parsed.lastRow(), parsed.firstColumn())
                .formatAsString(),
            new CellValue.Text("one"));
    sheet
        .cells()
        .setCell(
            new org.apache.poi.ss.util.CellReference(parsed.lastRow(), parsed.lastColumn())
                .formatAsString(),
            new CellValue.Text("two"));
  }

  private static void populateObservedTableRange(
      org.apache.poi.xssf.usermodel.XSSFSheet sheet, String range) {
    CellBounds parsed = CellBounds.parse(range);
    var header = sheet.getRow(parsed.firstRow());
    if (header == null) {
      header = sheet.createRow(parsed.firstRow());
    }
    header.createCell(parsed.firstColumn()).setCellValue("Header" + parsed.firstColumn());
    header.createCell(parsed.lastColumn()).setCellValue("Value" + parsed.lastColumn());
    var row = sheet.getRow(parsed.lastRow());
    if (row == null) {
      row = sheet.createRow(parsed.lastRow());
    }
    row.createCell(parsed.firstColumn()).setCellValue("one");
    row.createCell(parsed.lastColumn()).setCellValue("two");
  }

  private record UnicodeNameFixture(String namedRangeName, String tableName, String range) {}
}
