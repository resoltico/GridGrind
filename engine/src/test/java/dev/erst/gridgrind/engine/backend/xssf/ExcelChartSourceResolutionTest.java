package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.engine.charts.ChartAdapter;
import dev.erst.gridgrind.engine.charts.ChartAdapter.ResolvedChartSource;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.usermodel.FormulaError;
import org.junit.jupiter.api.Test;

/** Coverage for chart source resolution, cached scalar decoding, and failure messages. */
class ExcelChartSourceResolutionTest {
  @Test
  void sourceResolutionReadbacksAndFailuresStayDeterministic() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Charts");
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
                  new ChartTitleSpec.Formula("=B1")));
      ChartReport formulaTitleChart =
          ExcelChartTestSupport.chart(sheet.drawings().charts(), "FormulaTitle");
      ChartReport.Title.Formula chartTitle =
          assertInstanceOf(ChartReport.Title.Formula.class, formulaTitleChart.title());
      assertEquals("Plan", chartTitle.cachedText());
      assertInstanceOf(
          ChartReport.DataSource.NumericReference.class,
          ExcelChartTestSupport.singlePlot(formulaTitleChart, ChartReport.Line.class)
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
                  new ChartTitleSpec.None()));
      ChartReport sparseChart =
          ExcelChartTestSupport.chart(sheet.drawings().charts(), "SparseCategoriesChart");
      assertEquals(
          List.of("true", "", ""),
          assertInstanceOf(
                  ChartReport.DataSource.StringReference.class,
                  ExcelChartTestSupport.singlePlot(sparseChart, ChartReport.Line.class)
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
                  new ChartTitleSpec.None()));
      ChartReport formulaCategoriesChart =
          ExcelChartTestSupport.chart(sheet.drawings().charts(), "FormulaCategoriesChart");
      assertEquals(
          List.of("Alpha", "2.0", "true", ""),
          assertInstanceOf(
                  ChartReport.DataSource.StringReference.class,
                  ExcelChartTestSupport.singlePlot(formulaCategoriesChart, ChartReport.Line.class)
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
                          new ChartTitleSpec.None())));

      sheet
          .drawings()
          .setChart(
              chart(
                  "CrossSheetCategoriesChart",
                  ExcelChartTestSupport.ref("'Other'!A2:A4"),
                  ExcelChartTestSupport.ref("ChartActual"),
                  new ChartTitleSpec.None()));
      ChartReport crossSheetChart =
          ExcelChartTestSupport.chart(sheet.drawings().charts(), "CrossSheetCategoriesChart");
      assertEquals(
          List.of("North", "South", "West"),
          assertInstanceOf(
                  ChartReport.DataSource.StringReference.class,
                  ExcelChartTestSupport.singlePlot(crossSheetChart, ChartReport.Line.class)
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
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Charts");
      sheet.cells().setCell("A2", new CellValue.Text("Ari"));
      sheet.cells().setCell("A3", new CellValue.Text("Bo"));
      sheet.cells().setCell("A4", new CellValue.Text("Cy"));
      sheet.cells().setCell("B2", new ResolvedFormula("40+2"));
      sheet.cells().setCell("B3", new CellValue.NumberValue(7d));
      sheet.cells().setCell("B4", new ResolvedFormula("B3*2"));

      ResolvedChartSource resolved =
          ChartAdapter.resolveChartSource(
              sheet.xssfSheet(),
              "B2:B4",
              FormulaRuntime.poi(
                  workbook.xssfWorkbook().getCreationHelper().createFormulaEvaluator()));

      assertTrue(resolved.numeric());
      assertEquals(List.of(42d, 7d, 14d), resolved.numericValues());
      assertEquals(List.of("42.0", "7.0", "14.0"), resolved.stringValues());
    }
  }

  private static ResolvedChart chart(
      String name,
      ChartDataSourceSpec categories,
      ChartDataSourceSpec values,
      ResolvedChartTitle title) {
    return ExcelChartTestSupport.lineChart(
        name,
        ExcelChartTestSupport.anchor(4, 16, 10, 30),
        title,
        new ChartLegendSpec.Hidden(),
        ExcelChartDisplayBlanksAs.GAP,
        true,
        false,
        List.of(
            new ResolvedChart.Series(
                null,
                categories,
                values,
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty())));
  }

  private static void assertFailureMessage(
      XssfSheet sheet,
      String name,
      ChartDataSourceSpec categories,
      ChartDataSourceSpec values,
      String expectedMessagePart) {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                sheet
                    .drawings()
                    .setChart(chart(name, categories, values, new ChartTitleSpec.None())));
    assertTrue(failure.getMessage().contains(expectedMessagePart));
  }

  private static void seedChartSourceResolutionFixtures(XssfWorkbook workbook, XssfSheet sheet) {
    sheet.cells().setCell("D2", new CellValue.BooleanValue(true));
    sheet.cells().setCell("D3", new CellValue.Blank());
    workbook
        .names()
        .setNamedRange(
            new StructuredMutationAction.SetNamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName("NumericCategories"),
                new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Charts"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("B2:B4"))))
        .setNamedRange(
            new StructuredMutationAction.SetNamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName("SparseCategories"),
                new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Charts"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("D2:D4"))));

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
            new StructuredMutationAction.SetNamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName("FormulaCategories"),
                new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Charts"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("E2:E5"))))
        .setNamedRange(
            new StructuredMutationAction.SetNamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName("FormulaValues"),
                new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Charts"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("C2:C5"))))
        .setNamedRange(
            new StructuredMutationAction.SetNamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName("ErrorFormulaCategories"),
                new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Charts"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("F2:F4"))))
        .setNamedRange(
            new StructuredMutationAction.SetNamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName("ErrorCellCategories"),
                new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Charts"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("G2:G4"))));

    xssfSheet.getRow(1).createCell(5).setCellFormula("1/0");
    xssfSheet.getRow(2).createCell(5).setCellFormula("1/0");
    xssfSheet.getRow(3).createCell(5).setCellFormula("1/0");
    workbook.xssfWorkbook().getCreationHelper().createFormulaEvaluator().evaluateAll();
    xssfSheet.getRow(1).createCell(6).setCellErrorValue(FormulaError.NA.getCode());
    xssfSheet.getRow(2).createCell(6).setCellErrorValue(FormulaError.NA.getCode());
    xssfSheet.getRow(3).createCell(6).setCellErrorValue(FormulaError.NA.getCode());

    XssfSheet sourceSheet = workbook.getOrCreateSheet("Other");
    sourceSheet.cells().setCell("A2", new CellValue.Text("North"));
    sourceSheet.cells().setCell("A3", new CellValue.Text("South"));
    sourceSheet.cells().setCell("A4", new CellValue.Text("West"));
  }
}
