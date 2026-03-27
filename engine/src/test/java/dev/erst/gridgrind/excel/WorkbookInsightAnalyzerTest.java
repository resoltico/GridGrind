package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookInsightAnalyzer derived workbook analysis. */
class WorkbookInsightAnalyzerTest {
  @Test
  void executesEveryInsightCommandAgainstWorkbookState() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-insight-", ".xlsx");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var budget = workbook.createSheet("Budget");
      budget.createRow(0).createCell(0).setCellValue("Item");
      budget.getRow(0).createCell(1).setCellValue("Amount");
      budget.createRow(1).createCell(0).setCellValue("Hosting");
      budget.getRow(1).createCell(1).setCellValue(49.0);
      budget.createRow(2).createCell(0).setCellValue("Domain");
      budget.getRow(2).createCell(1).setCellValue(12.0);
      budget.createRow(3).createCell(1).setCellFormula("SUM(B2:B3)");
      var workbookName = workbook.createName();
      workbookName.setNameName("BudgetRollup");
      workbookName.setRefersToFormula("SUM(Budget!$B$2:$B$3)");
      var sheetName = workbook.createName();
      sheetName.setNameName("LocalItem");
      sheetName.setSheetIndex(0);
      sheetName.setRefersToFormula("Budget!$A$1");
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      WorkbookInsightAnalyzer analyzer = new WorkbookInsightAnalyzer();

      WorkbookReadResult.FormulaSurfaceResult formulaSurface =
          cast(
              WorkbookReadResult.FormulaSurfaceResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookReadCommand.AnalyzeFormulaSurface(
                      "formula", new ExcelSheetSelection.All())));
      WorkbookReadResult.SheetSchemaResult sheetSchema =
          cast(
              WorkbookReadResult.SheetSchemaResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookReadCommand.AnalyzeSheetSchema("schema", "Budget", "A1", 4, 2)));
      WorkbookReadResult.NamedRangeSurfaceResult namedRangeSurface =
          cast(
              WorkbookReadResult.NamedRangeSurfaceResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookReadCommand.AnalyzeNamedRangeSurface(
                      "ranges", new ExcelNamedRangeSelection.All())));

      assertEquals(1, formulaSurface.analysis().totalFormulaCellCount());
      assertEquals("Budget", formulaSurface.analysis().sheets().getFirst().sheetName());
      assertEquals(
          "SUM(B2:B3)",
          formulaSurface.analysis().sheets().getFirst().formulas().getFirst().formula());
      assertEquals(
          List.of("B4"),
          formulaSurface.analysis().sheets().getFirst().formulas().getFirst().addresses());

      assertEquals("Budget", sheetSchema.analysis().sheetName());
      assertEquals("A1", sheetSchema.analysis().topLeftAddress());
      assertEquals(3, sheetSchema.analysis().dataRowCount());
      assertEquals("STRING", sheetSchema.analysis().columns().getFirst().dominantType());

      assertEquals(1, namedRangeSurface.analysis().workbookScopedCount());
      assertEquals(1, namedRangeSurface.analysis().sheetScopedCount());
      assertEquals(1, namedRangeSurface.analysis().rangeBackedCount());
      assertEquals(1, namedRangeSurface.analysis().formulaBackedCount());
    }
  }

  @Test
  void analyzeSheetSchemaReturnsNullDominantTypeOnTies() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Budget");
      sheet.setCell("A1", ExcelCellValue.text("Mixed"));
      sheet.setCell("A2", ExcelCellValue.text("text"));
      sheet.setCell("A3", ExcelCellValue.number(1.0));

      WorkbookReadResult.SheetSchema schema =
          new WorkbookInsightAnalyzer().analyzeSheetSchema(workbook, "Budget", "A1", 3, 1);

      assertNull(schema.columns().getFirst().dominantType());
    }
  }

  @Test
  void analyzesSelectedSheetsAndNamedRanges() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      budget.setCell("A1", ExcelCellValue.text("Item"));
      budget.setCell("B1", ExcelCellValue.text("Amount"));
      budget.setCell("B2", ExcelCellValue.formula("1+1"));

      ExcelSheet forecast = workbook.getOrCreateSheet("Forecast");
      forecast.setCell("A1", ExcelCellValue.text("Item"));
      forecast.setCell("B1", ExcelCellValue.text("Amount"));
      forecast.setCell("B2", ExcelCellValue.formula("2+2"));

      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "BudgetTotal",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Budget", "B2")));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "LocalForecast",
              new ExcelNamedRangeScope.SheetScope("Forecast"),
              new ExcelNamedRangeTarget("Forecast", "B2")));

      WorkbookInsightAnalyzer analyzer = new WorkbookInsightAnalyzer();

      WorkbookReadResult.FormulaSurface formulaSurface =
          analyzer.analyzeFormulaSurface(
              workbook, new ExcelSheetSelection.Selected(List.of("Forecast")));
      assertEquals(1, formulaSurface.totalFormulaCellCount());
      assertEquals(
          List.of("Forecast"),
          formulaSurface.sheets().stream()
              .map(WorkbookReadResult.SheetFormulaSurface::sheetName)
              .toList());

      WorkbookReadResult.NamedRangeSurface namedRangeSurface =
          analyzer.analyzeNamedRangeSurface(
              workbook,
              new ExcelNamedRangeSelection.Selected(
                  List.of(
                      new ExcelNamedRangeSelector.WorkbookScope("BudgetTotal"),
                      new ExcelNamedRangeSelector.SheetScope("LocalForecast", "Forecast"))));
      assertEquals(1, namedRangeSurface.workbookScopedCount());
      assertEquals(1, namedRangeSurface.sheetScopedCount());
      assertEquals(2, namedRangeSurface.rangeBackedCount());
      assertEquals(0, namedRangeSurface.formulaBackedCount());
    }
  }

  @Test
  void rejectsNullWorkbookAndSelections() {
    WorkbookInsightAnalyzer analyzer = new WorkbookInsightAnalyzer();

    assertThrows(
        NullPointerException.class,
        () ->
            analyzer.execute(
                null,
                new WorkbookReadCommand.AnalyzeFormulaSurface(
                    "formula", new ExcelSheetSelection.All())));
    assertThrows(NullPointerException.class, () -> analyzer.execute(ExcelWorkbook.create(), null));
    assertThrows(
        NullPointerException.class,
        () -> analyzer.analyzeFormulaSurface(null, new ExcelSheetSelection.All()));
    assertThrows(
        NullPointerException.class,
        () -> analyzer.analyzeFormulaSurface(ExcelWorkbook.create(), null));
    assertThrows(
        NullPointerException.class,
        () -> analyzer.analyzeNamedRangeSurface(null, new ExcelNamedRangeSelection.All()));
    assertThrows(
        NullPointerException.class,
        () -> analyzer.analyzeNamedRangeSurface(ExcelWorkbook.create(), null));
  }

  private static <T> T cast(Class<T> type, Object value) {
    return type.cast(assertInstanceOf(type, value));
  }
}
