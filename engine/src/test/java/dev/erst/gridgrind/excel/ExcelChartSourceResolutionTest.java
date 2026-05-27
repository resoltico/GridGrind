package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.usermodel.FormulaError;
import org.junit.jupiter.api.Test;

/** Coverage for chart source resolution, cached scalar decoding, and failure messages. */
class ExcelChartSourceResolutionTest {
  @Test
  void sourceResolutionReadbacksAndFailuresStayDeterministic() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      ExcelChartTestSupport.seedChartNamedRanges(workbook, "Charts");
      seedChartSourceResolutionFixtures(workbook, sheet);

      sheet
          .drawings()
          .setChart(
              chart(
                  "FormulaTitle",
                  ExcelChartTestSupport.ref("NumericCategories"),
                  ExcelChartTestSupport.ref("ChartActual"),
                  new ExcelChartDefinition.Title.Formula("=B1")));
      ExcelChartSnapshot formulaTitleChart =
          ExcelChartTestSupport.chart(sheet.drawings().charts(), "FormulaTitle");
      ExcelChartSnapshot.Title.Formula chartTitle =
          assertInstanceOf(ExcelChartSnapshot.Title.Formula.class, formulaTitleChart.title());
      assertEquals("Plan", chartTitle.cachedText());
      assertInstanceOf(
          ExcelChartSnapshot.DataSource.NumericReference.class,
          ExcelChartTestSupport.singlePlot(formulaTitleChart, ExcelChartSnapshot.Line.class)
              .series()
              .getFirst()
              .categories());

      sheet
          .drawings()
          .setChart(
              chart(
                  "SparseCategoriesChart",
                  ExcelChartTestSupport.ref("SparseCategories"),
                  ExcelChartTestSupport.ref("ChartActual"),
                  new ExcelChartDefinition.Title.None()));
      ExcelChartSnapshot sparseChart =
          ExcelChartTestSupport.chart(sheet.drawings().charts(), "SparseCategoriesChart");
      assertEquals(
          List.of("true", "", ""),
          assertInstanceOf(
                  ExcelChartSnapshot.DataSource.StringReference.class,
                  ExcelChartTestSupport.singlePlot(sparseChart, ExcelChartSnapshot.Line.class)
                      .series()
                      .getFirst()
                      .categories())
              .cachedValues());

      sheet
          .drawings()
          .setChart(
              chart(
                  "FormulaCategoriesChart",
                  ExcelChartTestSupport.ref("FormulaCategories"),
                  ExcelChartTestSupport.ref("FormulaValues"),
                  new ExcelChartDefinition.Title.None()));
      ExcelChartSnapshot formulaCategoriesChart =
          ExcelChartTestSupport.chart(sheet.drawings().charts(), "FormulaCategoriesChart");
      assertEquals(
          List.of("Alpha", "2.0", "true", ""),
          assertInstanceOf(
                  ExcelChartSnapshot.DataSource.StringReference.class,
                  ExcelChartTestSupport.singlePlot(
                          formulaCategoriesChart, ExcelChartSnapshot.Line.class)
                      .series()
                      .getFirst()
                      .categories())
              .cachedValues());

      assertDoesNotThrow(
          () ->
              sheet
                  .drawings()
                  .setChart(
                      chart(
                          "LocalCategoriesChart",
                          ExcelChartTestSupport.ref("LocalCategories"),
                          ExcelChartTestSupport.ref("ChartActual"),
                          new ExcelChartDefinition.Title.None())));

      sheet
          .drawings()
          .setChart(
              chart(
                  "CrossSheetCategoriesChart",
                  ExcelChartTestSupport.ref("'Other'!A2:A4"),
                  ExcelChartTestSupport.ref("ChartActual"),
                  new ExcelChartDefinition.Title.None()));
      ExcelChartSnapshot crossSheetChart =
          ExcelChartTestSupport.chart(sheet.drawings().charts(), "CrossSheetCategoriesChart");
      assertEquals(
          List.of("North", "South", "West"),
          assertInstanceOf(
                  ExcelChartSnapshot.DataSource.StringReference.class,
                  ExcelChartTestSupport.singlePlot(crossSheetChart, ExcelChartSnapshot.Line.class)
                      .series()
                      .getFirst()
                      .categories())
              .cachedValues());

      assertFailureMessage(
          sheet,
          "StringValues",
          ExcelChartTestSupport.ref("ChartCategories"),
          ExcelChartTestSupport.ref("LocalCategories"),
          "Chart value source must resolve to numeric cells");
      assertFailureMessage(
          sheet,
          "NonContiguous",
          ExcelChartTestSupport.ref("A2:A4,C2:C4"),
          ExcelChartTestSupport.ref("ChartActual"),
          "one contiguous area");
      assertFailureMessage(
          sheet,
          "InvalidCellReference",
          ExcelChartTestSupport.ref("A2:???"),
          ExcelChartTestSupport.ref("ChartActual"),
          "one contiguous area");
      assertFailureMessage(
          sheet,
          "MultiAreaDefinedName",
          ExcelChartTestSupport.ref("MultiAreaSource"),
          ExcelChartTestSupport.ref("ChartActual"),
          "must resolve to one contiguous area");
      assertFailureMessage(
          sheet,
          "MissingSheet",
          ExcelChartTestSupport.ref("'Missing'!A2:A4"),
          ExcelChartTestSupport.ref("ChartActual"),
          "resolves to missing sheet");
      assertFailureMessage(
          sheet,
          "ErrorFormulaCategoriesChart",
          ExcelChartTestSupport.ref("ErrorFormulaCategories"),
          ExcelChartTestSupport.ref("ChartActual"),
          "must not cache error values");
      assertFailureMessage(
          sheet,
          "ErrorCellCategoriesChart",
          ExcelChartTestSupport.ref("ErrorCellCategories"),
          ExcelChartTestSupport.ref("ChartActual"),
          "must not contain error values");
    }
  }

  @Test
  void runtimeAwareSourceResolutionEvaluatesFormulaCellsBeforeReadingValues() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Charts");
      sheet.cells().setCell("A2", ExcelCellValue.text("Ari"));
      sheet.cells().setCell("A3", ExcelCellValue.text("Bo"));
      sheet.cells().setCell("A4", ExcelCellValue.text("Cy"));
      sheet.cells().setCell("B2", ExcelCellValue.formula("40+2"));
      sheet.cells().setCell("B3", ExcelCellValue.number(7d));
      sheet.cells().setCell("B4", ExcelCellValue.formula("B3*2"));

      ResolvedChartSource resolved =
          ExcelChartSourceSupport.resolveChartSource(
              sheet.xssfSheet(),
              "B2:B4",
              ExcelFormulaRuntime.poi(
                  workbook.xssfWorkbook().getCreationHelper().createFormulaEvaluator()));

      assertTrue(resolved.numeric());
      assertEquals(List.of(42d, 7d, 14d), resolved.numericValues());
      assertEquals(List.of("42.0", "7.0", "14.0"), resolved.stringValues());
    }
  }

  private static ExcelChartDefinition chart(
      String name,
      ExcelChartDefinition.DataSource categories,
      ExcelChartDefinition.DataSource values,
      ExcelChartDefinition.Title title) {
    return ExcelChartTestSupport.lineChart(
        name,
        ExcelChartTestSupport.anchor(4, 16, 10, 30),
        title,
        new ExcelChartDefinition.Legend.Hidden(),
        ExcelChartDisplayBlanksAs.GAP,
        true,
        false,
        List.of(
            new ExcelChartDefinition.Series(
                null,
                categories,
                values,
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty())));
  }

  private static void assertFailureMessage(
      ExcelSheet sheet,
      String name,
      ExcelChartDefinition.DataSource categories,
      ExcelChartDefinition.DataSource values,
      String expectedMessagePart) {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                sheet
                    .drawings()
                    .setChart(
                        chart(name, categories, values, new ExcelChartDefinition.Title.None())));
    assertTrue(failure.getMessage().contains(expectedMessagePart));
  }

  private static void seedChartSourceResolutionFixtures(ExcelWorkbook workbook, ExcelSheet sheet) {
    sheet.cells().setCell("D2", ExcelCellValue.bool(true));
    sheet.cells().setCell("D3", ExcelCellValue.blank());
    workbook
        .names()
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "NumericCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range("Charts", "B2:B4")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "SparseCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range("Charts", "D2:D4")));

    var localCategories = workbook.xssfWorkbook().createName();
    localCategories.setNameName("LocalCategories");
    localCategories.setSheetIndex(workbook.xssfWorkbook().getSheetIndex("Charts"));
    localCategories.setRefersToFormula("A2:A4");

    var multiAreaSource = workbook.xssfWorkbook().createName();
    multiAreaSource.setNameName("MultiAreaSource");
    multiAreaSource.setRefersToFormula("A2:A4,C2:C4");

    var xssfSheet = sheet.xssfSheet();
    xssfSheet.getRow(1).createCell(4).setCellFormula("\"Alpha\"");
    xssfSheet.getRow(2).createCell(4).setCellFormula("1+1");
    xssfSheet.getRow(3).createCell(4).setCellFormula("TRUE");
    xssfSheet.createRow(4).createCell(4).setCellFormula("\"\"");
    xssfSheet.getRow(4).createCell(2).setCellValue(25d);
    workbook.xssfWorkbook().getCreationHelper().createFormulaEvaluator().evaluateAll();

    workbook
        .names()
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "FormulaCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range("Charts", "E2:E5")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "FormulaValues",
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range("Charts", "C2:C5")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ErrorFormulaCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range("Charts", "F2:F4")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ErrorCellCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range("Charts", "G2:G4")));

    xssfSheet.getRow(1).createCell(5).setCellFormula("1/0");
    xssfSheet.getRow(2).createCell(5).setCellFormula("1/0");
    xssfSheet.getRow(3).createCell(5).setCellFormula("1/0");
    workbook.xssfWorkbook().getCreationHelper().createFormulaEvaluator().evaluateAll();
    xssfSheet.getRow(1).createCell(6).setCellErrorValue(FormulaError.NA.getCode());
    xssfSheet.getRow(2).createCell(6).setCellErrorValue(FormulaError.NA.getCode());
    xssfSheet.getRow(3).createCell(6).setCellErrorValue(FormulaError.NA.getCode());

    ExcelSheet sourceSheet = workbook.getOrCreateSheet("Other");
    sourceSheet.cells().setCell("A2", ExcelCellValue.text("North"));
    sourceSheet.cells().setCell("A3", ExcelCellValue.text("South"));
    sourceSheet.cells().setCell("A4", ExcelCellValue.text("West"));
  }
}
