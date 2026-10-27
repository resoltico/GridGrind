package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.drawing.CommentAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintMarginsSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintSetupSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.style.ResolvedRichTextRun;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableColumnSpec;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.autofilter.AutofilterAdapter;
import dev.erst.gridgrind.engine.conditionalformatting.ConditionalFormattingAdapter;
import dev.erst.gridgrind.engine.formulas.ComparisonFormulas;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationType;

/** Tests for sheet-copy planning and validation helpers. */
class ExcelSheetCopyControllerTest {
  @Test
  void copiesNativeHiddenZeroHeightRows(@org.junit.jupiter.api.io.TempDir Path directory)
      throws IOException {
    Path artifact = directory.resolve("copied.xlsx");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Source");
      var row = workbook.xssfWorkbook().getSheet("Source").createRow(0);
      row.setHeight((short) 0);
      row.setZeroHeight(true);
      row.createCell(0).setCellValue("preserved");
      new SheetCopyAdapter()
          .copySheet(workbook, "Source", "Copied", new SheetCopyPosition.AppendAtEnd());
      workbook
          .persistence()
          .save(
              artifact,
              WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }
    try (var input = Files.newInputStream(artifact);
        XSSFWorkbook saved = new XSSFWorkbook(input)) {
      for (String name : List.of("Source", "Copied")) {
        var row = saved.getSheet(name).getRow(0);
        assertEquals(0, row.getHeight());
        assertTrue(row.getZeroHeight());
        assertEquals("preserved", row.getCell(0).getStringCellValue());
      }
    }
  }

  @Test
  void copyPreservesObservedStylelessRuleAndRetargetsFormula() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Source");
      var source = workbook.xssfWorkbook().getSheet("Source");
      var formatting = source.getSheetConditionalFormatting();
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("A1:A3")},
          formatting.createConditionalFormattingRule("Source!A1>0"));
      new SheetCopyAdapter()
          .copySheet(workbook, "Source", "Copied", new SheetCopyPosition.AppendAtEnd());
      var copied =
          workbook
              .xssfWorkbook()
              .getSheet("Copied")
              .getSheetConditionalFormatting()
              .getConditionalFormattingAt(0)
              .getRule(0);
      assertEquals("Copied!A1>0", copied.getFormula1());
      assertFalse(copied.getStopIfTrue());
      assertNull(copied.getFontFormatting());
      assertEquals(
          "Source!A1>0", formatting.getConditionalFormattingAt(0).getRule(0).getFormula1());
    }
  }

  @Test
  void copyRejectsUnsupportedStyleBeforeCloning() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Source");
      var source = workbook.xssfWorkbook().getSheet("Source");
      var formatting = source.getSheetConditionalFormatting();
      var rule = formatting.createConditionalFormattingRule("A1>0");
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("A1:A3")}, rule);
      var style = org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf.Factory.newInstance();
      style.addNewFont().addNewName().setVal("Unsupported font family");
      ConditionalFormattingAdapter.attachStyle(
          workbook.xssfWorkbook().getStylesSource(),
          source.getCTWorksheet().getConditionalFormattingArray(0).getCfRuleArray(0),
          style);
      assertThrows(
          IllegalArgumentException.class,
          () ->
              new SheetCopyAdapter()
                  .copySheet(workbook, "Source", "Copied", new SheetCopyPosition.AppendAtEnd()));
      assertEquals(1, workbook.xssfWorkbook().getNumberOfSheets());
    }
  }

  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC");

  @Test
  void copySheetPreservesProtectionOnEmptySourceSheets() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Source");
      workbook.protection().setSheetProtection("Source", protectionSettings());

      workbook.sheets().copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd());

      assertEquals(List.of("Source", "Replica"), workbook.sheets().sheetNames());
      assertEquals(
          new SheetProtectionReport.Protected(protectionSettings()),
          workbook.sheetSummary("Replica").protection());
      assertEquals(0, workbook.sheetSummary("Replica").physicalRowCount());
    }
  }

  @Test
  void copySheetSupportsReopenedPoiCommentWorkbooks() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-comments-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Source");
      sheet.createRow(0).createCell(0).setCellValue("Lead");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFClientAnchor anchor = drawing.createAnchor(64, 24, 448, 96, 0, 0, 3, 3);
      XSSFComment comment = drawing.createCellComment(anchor);
      comment.setString(new XSSFRichTextString("Review"));
      comment.setAuthor("GridGrind");
      sheet.getRow(0).getCell(0).setCellComment(comment);
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (XssfWorkbook workbook =
        XssfWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertDoesNotThrow(
          () ->
              workbook
                  .sheets()
                  .copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd()));
      assertEquals(
          "Review", workbook.sheet("Replica").comments().comments().getFirst().comment().text());
    }
  }

  @Test
  void copySheetPreservesSheetProtectionPasswordHashes() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Source");
      workbook
          .protection()
          .setSheetProtection("Source", protectionSettings(), Optional.of("gridgrind-copy"));

      workbook.sheets().copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd());

      assertTrue(
          workbook.xssfWorkbook().getSheet("Replica").validateSheetPassword("gridgrind-copy"));
      assertFalse(
          workbook.xssfWorkbook().getSheet("Replica").validateSheetPassword("gridgrind-wrong"));
      assertEquals(
          new SheetProtectionReport.Protected(protectionSettings()),
          workbook.sheetSummary("Replica").protection());
    }
  }

  @Test
  void copySheetPreservesAdvancedPrintSetupWhenOnlyPageSetupAttrsAreNonDefault()
      throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Source");
      ResolvedPrintLayout advancedPrintLayout =
          new ResolvedPrintLayout(
              new PrintAreaSpec.None(),
              ExcelPrintOrientation.PORTRAIT,
              new PrintScalingSpec.Automatic(),
              new PrintTitleRowsSpec.None(),
              new PrintTitleColumnsSpec.None(),
              HeaderFooterTextReport.blank(),
              HeaderFooterTextReport.blank(),
              new PrintSetupSpec(
                  new PrintMarginsSpec(0.35d, 0.55d, 0.6d, 0.45d, 0.3d, 0.3d),
                  false,
                  true,
                  true,
                  8,
                  true,
                  true,
                  2,
                  true,
                  4,
                  List.of(6),
                  List.of(3)));
      workbook.sheet("Source").layout().setPrintLayout(advancedPrintLayout);

      workbook.sheets().copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd());

      assertEquals(advancedPrintLayout, workbook.sheet("Replica").layout().printLayout());
    }
  }

  @Test
  void copySheetPreservesDrawingObjectsAndRetargetsCopiedCharts() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-drawings-");

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet source = workbook.getOrCreateSheet("Source");
      ExcelChartTestSupport.seedChartData(source);
      source
          .drawings()
          .setPicture(
              new ResolvedPicture(
                  "OpsPicture",
                  new BinaryContent(PNG_PIXEL_BYTES),
                  ExcelPictureFormat.PNG,
                  ExcelChartTestSupport.anchor(0, 5, 3, 9),
                  Optional.of("Queue preview")));
      source
          .drawings()
          .setChart(
              ExcelChartTestSupport.barChart(
                  "OpsChart",
                  ExcelChartTestSupport.anchor(6, 1, 12, 14),
                  new ResolvedChartTitle.Text("Roadmap"),
                  new ChartLegendSpec.Hidden(),
                  ExcelChartDisplayBlanksAs.GAP,
                  true,
                  false,
                  ExcelChartBarDirection.COLUMN,
                  List.of(
                      new ResolvedChart.Series(
                          new ResolvedChartTitle.Text("Plan"),
                          ExcelChartTestSupport.ref("A2:A4"),
                          ExcelChartTestSupport.ref("B2:B4")))));

      workbook.sheets().copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd());

      XssfSheet replica = workbook.sheet("Replica");
      assertEquals(
          List.of("OpsPicture", "OpsChart"),
          replica.drawings().drawingObjects().stream().map(DrawingObjectReport::name).toList());
      DrawingObjectReport.Picture copiedPicture =
          assertInstanceOf(
              DrawingObjectReport.Picture.class,
              replica.drawings().drawingObjects().stream()
                  .filter(snapshot -> "OpsPicture".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(
          DrawingGeometryFixtures.report(ExcelChartTestSupport.anchor(0, 5, 3, 9)),
          copiedPicture.anchor());

      ChartReport copiedChart = replica.drawings().charts().getFirst();
      assertEquals("OpsChart", copiedChart.name());
      assertEquals(
          DrawingGeometryFixtures.report(ExcelChartTestSupport.anchor(6, 1, 12, 14)),
          copiedChart.anchor());
      ChartReport.Series copiedSeries =
          ExcelChartTestSupport.singlePlot(copiedChart, ChartReport.Bar.class).series().getFirst();
      assertEquals(
          "Replica!$A$2:$A$4",
          assertInstanceOf(ChartReport.DataSource.StringReference.class, copiedSeries.categories())
              .formula());
      assertEquals(
          "Replica!$B$2:$B$4",
          assertInstanceOf(ChartReport.DataSource.NumericReference.class, copiedSeries.values())
              .formula());

      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XssfWorkbook reopened =
        XssfWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      XssfSheet replica = reopened.sheet("Replica");
      assertEquals(
          List.of("OpsPicture", "OpsChart"),
          replica.drawings().drawingObjects().stream().map(DrawingObjectReport::name).toList());
      ChartReport copiedChart = replica.drawings().charts().getFirst();
      ChartReport.Series copiedSeries =
          ExcelChartTestSupport.singlePlot(copiedChart, ChartReport.Bar.class).series().getFirst();
      assertEquals(
          "Replica!$A$2:$A$4",
          assertInstanceOf(ChartReport.DataSource.StringReference.class, copiedSeries.categories())
              .formula());
    }
  }

  @Test
  void copySheetPreservesMultipleChartsWithoutFramelessRelations() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet source = workbook.getOrCreateSheet("Source");
      ExcelChartTestSupport.seedChartData(source);
      source
          .drawings()
          .setChart(
              ExcelChartTestSupport.barChart(
                  "OpsBar",
                  ExcelChartTestSupport.anchor(0, 5, 6, 16),
                  new ResolvedChartTitle.Text("Plan"),
                  new ChartLegendSpec.Hidden(),
                  ExcelChartDisplayBlanksAs.GAP,
                  true,
                  false,
                  ExcelChartBarDirection.COLUMN,
                  List.of(
                      new ResolvedChart.Series(
                          new ResolvedChartTitle.Text("Plan"),
                          ExcelChartTestSupport.ref("A2:A4"),
                          ExcelChartTestSupport.ref("B2:B4")))));
      source
          .drawings()
          .setChart(
              ExcelChartTestSupport.lineChart(
                  "OpsLine",
                  ExcelChartTestSupport.anchor(7, 5, 13, 16),
                  new ChartTitleSpec.Formula("C1"),
                  new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT),
                  ExcelChartDisplayBlanksAs.SPAN,
                  true,
                  false,
                  List.of(
                      new ResolvedChart.Series(
                          new ResolvedChartTitle.Text("Actual"),
                          ExcelChartTestSupport.ref("A2:A4"),
                          ExcelChartTestSupport.ref("C2:C4")))));

      workbook.sheets().copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd());

      XssfSheet replica = workbook.sheet("Replica");
      assertEquals(
          List.of("OpsBar", "OpsLine"),
          replica.drawings().charts().stream().map(ChartReport::name).toList());
      assertEquals(
          List.of("OpsBar", "OpsLine"),
          replica.drawings().drawingObjects().stream().map(DrawingObjectReport::name).toList());

      XSSFDrawing drawing = replica.xssfSheet().getDrawingPatriarch();
      assertNotNull(drawing);
      drawing.getShapes();
      assertTrue(drawing.getCharts().stream().allMatch(chart -> chart.getGraphicFrame() != null));
    }
  }

  @Test
  void copySheetSupportsChartsBackedByWorkbookNamedRanges() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-chart-names-");

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet source = workbook.getOrCreateSheet("Source");
      ExcelChartTestSupport.seedChartData(source);
      ExcelChartTestSupport.seedChartNamedRanges(workbook, "Source");
      source
          .drawings()
          .setChart(
              ExcelChartTestSupport.barChart(
                  "NamedRangeChart",
                  ExcelChartTestSupport.anchor(4, 1, 10, 14),
                  new ResolvedChartTitle.Text("Named ranges"),
                  new ChartLegendSpec.Hidden(),
                  ExcelChartDisplayBlanksAs.GAP,
                  true,
                  false,
                  ExcelChartBarDirection.COLUMN,
                  List.of(
                      new ResolvedChart.Series(
                          new ResolvedChartTitle.Text("Plan"),
                          ExcelChartTestSupport.ref("ChartCategories"),
                          ExcelChartTestSupport.ref("ChartPlan")))));

      assertDoesNotThrow(
          () ->
              workbook
                  .sheets()
                  .copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd()));
      assertEquals(List.of("Source", "Replica"), workbook.sheets().sheetNames());

      ChartReport sourceChart = workbook.sheet("Source").drawings().charts().getFirst();
      ChartReport.Series sourceSeries =
          ExcelChartTestSupport.singlePlot(sourceChart, ChartReport.Bar.class).series().getFirst();
      assertEquals(
          "ChartCategories",
          assertInstanceOf(ChartReport.DataSource.StringReference.class, sourceSeries.categories())
              .formula());
      assertEquals(
          "ChartPlan",
          assertInstanceOf(ChartReport.DataSource.NumericReference.class, sourceSeries.values())
              .formula());

      ChartReport copiedChart = workbook.sheet("Replica").drawings().charts().getFirst();
      ChartReport.Series copiedSeries =
          ExcelChartTestSupport.singlePlot(copiedChart, ChartReport.Bar.class).series().getFirst();
      assertEquals(
          "Replica!$A$2:$A$4",
          assertInstanceOf(ChartReport.DataSource.StringReference.class, copiedSeries.categories())
              .formula());
      assertEquals(
          "Replica!$B$2:$B$4",
          assertInstanceOf(ChartReport.DataSource.NumericReference.class, copiedSeries.values())
              .formula());

      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XssfWorkbook reopened =
        XssfWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ChartReport sourceChart = reopened.sheet("Source").drawings().charts().getFirst();
      ChartReport.Series sourceSeries =
          ExcelChartTestSupport.singlePlot(sourceChart, ChartReport.Bar.class).series().getFirst();
      assertEquals(
          "ChartCategories",
          assertInstanceOf(ChartReport.DataSource.StringReference.class, sourceSeries.categories())
              .formula());
      assertEquals(
          "ChartPlan",
          assertInstanceOf(ChartReport.DataSource.NumericReference.class, sourceSeries.values())
              .formula());

      ChartReport copiedChart = reopened.sheet("Replica").drawings().charts().getFirst();
      ChartReport.Series copiedSeries =
          ExcelChartTestSupport.singlePlot(copiedChart, ChartReport.Bar.class).series().getFirst();
      assertEquals(
          "Replica!$A$2:$A$4",
          assertInstanceOf(ChartReport.DataSource.StringReference.class, copiedSeries.categories())
              .formula());
      assertEquals(
          "Replica!$B$2:$B$4",
          assertInstanceOf(ChartReport.DataSource.NumericReference.class, copiedSeries.values())
              .formula());
    }
  }

  @Test
  void copySheetRetargetsAdvancedWorkbookCoreStructures() throws IOException {
    AutofilterAdapter autofilterAdapter = new AutofilterAdapter();
    TableAdapter tableController = new TableAdapter();

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet source = seedAdvancedCopySource(workbook);
      List<CellCommentReport> sourceComments =
          source
              .comments()
              .comments(
                  List.of("E2").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.CellAddress::new)
                      .toList());

      workbook.sheets().copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd());

      XssfSheet replica = workbook.sheet("Replica");
      assertEquals("SUM(Replica!B2:B3)", replica.cells().formula("C2"));
      assertEquals(
          sourceComments,
          replica
              .comments()
              .comments(
                  List.of("E2").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.CellAddress::new)
                      .toList()));
      assertEquals(
          List.of(
              new AutofilterEntryReport.SheetOwned(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:F3"),
                  List.of(
                      new AutofilterFilterColumnReport(
                          0L,
                          false,
                          new AutofilterFilterCriterionReport.Values(List.of("Ada"), true)),
                      new AutofilterFilterColumnReport(
                          1L,
                          true,
                          new AutofilterFilterCriterionReport.Custom(
                              true,
                              List.of(
                                  new AutofilterFilterCriterionReport.CustomConditionReport(
                                      "greaterThan", "1"),
                                  new AutofilterFilterCriterionReport.CustomConditionReport(
                                      "equal", "Queue")))),
                      new AutofilterFilterColumnReport(
                          2L,
                          true,
                          new AutofilterFilterCriterionReport.Dynamic("today", 1.0d, 2.0d)),
                      new AutofilterFilterColumnReport(
                          3L,
                          true,
                          new AutofilterFilterCriterionReport.Top10(false, true, 10.0d, null)),
                      new AutofilterFilterColumnReport(
                          4L,
                          true,
                          new AutofilterFilterCriterionReport.Color(
                              false, CellColorReport.rgb(new RgbHex("#AABBCC")))),
                      new AutofilterFilterColumnReport(
                          5L,
                          true,
                          new AutofilterFilterCriterionReport.Icon("3TrafficLights1", 2))),
                  java.util.Optional.of(
                      new AutofilterSortStateReport(
                          new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:F3"),
                          true,
                          true,
                          java.util.Optional.of(
                              dev.erst.gridgrind.core.vocabulary.ExcelAutofilterSortMethod.STROKE),
                          List.of(
                              new AutofilterSortConditionReport.CellColor(
                                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A2:A3"),
                                  true,
                                  CellColorReport.rgb(new RgbHex("#102030"))),
                              new AutofilterSortConditionReport.Icon(
                                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("B2:B3"),
                                  false,
                                  4)))))),
          autofilterAdapter.sheetOwnedAutofilters(replica.xssfSheet()));
      List<DataValidationEntryReport> replicaValidations = replica.metadata().dataValidations();
      assertEquals(2, replicaValidations.size());
      DataValidationEntryReport.Supported formulaList =
          assertInstanceOf(DataValidationEntryReport.Supported.class, replicaValidations.get(0));
      assertEquals(
          new DataValidationRuleSpec.FormulaList(
              ComparisonFormulas.normalizeFormula("Replica!$J$2:$J$3", "formula")),
          formulaList.validation().rule());
      DataValidationEntryReport.Supported wholeNumber =
          assertInstanceOf(DataValidationEntryReport.Supported.class, replicaValidations.get(1));
      assertInstanceOf(DataValidationRuleSpec.WholeNumber.class, wholeNumber.validation().rule());
      List<CTDataValidation> rawReplicaValidations =
          List.of(
              replica.xssfSheet().getCTWorksheet().getDataValidations().getDataValidationArray());
      assertEquals("Replica!$J$2:$J$3", rawReplicaValidations.get(0).getFormula1());
      assertEquals("1", rawReplicaValidations.get(1).getFormula1());
      assertTrue(rawReplicaValidations.get(1).getFormula2().contains("Replica"));
      assertTrue(rawReplicaValidations.get(1).getFormula2().contains("$B$2:$B$3"));
      List<NamedRangeReport> replicaNames =
          workbook.names().namedRanges().stream()
              .filter(
                  namedRange ->
                      namedRange.scope() instanceof NamedRangeScopeReport.Sheet scope
                          && "Replica".equals(scope.sheetName().value()))
              .toList();
      assertEquals(2, replicaNames.size());
      assertTrue(
          replicaNames.stream()
              .anyMatch(
                  namedRange ->
                      namedRange instanceof NamedRangeReport.RangeReport rangeSnapshot
                          && "LocalRange".equals(rangeSnapshot.name().value())
                          && rangeSnapshot.refersToFormula().contains("Replica")
                          && rangeSnapshot.refersToFormula().contains("$A$2")
                          && rangeSnapshot.refersToFormula().contains("$A$3")
                          && "A2:A3"
                              .equals(
                                  assertInstanceOf(
                                          NamedRangeTarget.Range.class, rangeSnapshot.target())
                                      .range()
                                      .value())));
      assertTrue(
          replicaNames.stream()
              .anyMatch(
                  namedRange ->
                      namedRange instanceof NamedRangeReport.FormulaReport formulaSnapshot
                          && "LocalFormula".equals(formulaSnapshot.name().value())
                          && formulaSnapshot.refersToFormula().contains("Replica")
                          && formulaSnapshot.refersToFormula().contains("$B$2:$B$3")));
      ConditionalFormattingEntryReport copiedFormatting =
          replica.metadata().conditionalFormatting().getFirst();
      assertEquals(
          "SUM(Replica!$B$2:$B$3)>0",
          assertInstanceOf(
                  ConditionalFormattingRuleReport.FormulaRule.class,
                  copiedFormatting.rules().get(0))
              .formula());
      assertEquals(
          "SUM(Replica!$B$2:$B$3)",
          assertInstanceOf(
                  ConditionalFormattingRuleReport.CellValueRule.class,
                  copiedFormatting.rules().get(1))
              .formula2()
              .orElseThrow());
      TableEntryReport copiedTable =
          tableController.tables(workbook, new TableSelector.All()).stream()
              .filter(table -> "Replica".equals(table.sheetName().value()))
              .findFirst()
              .orElseThrow();
      assertEquals("L1:M3", copiedTable.range().value());
      assertNotEquals("ReplicaTable", copiedTable.name());
      assertEquals("replica table", copiedTable.presentation().comment().orElseThrow());
      assertTrue(copiedTable.behavior().published());
      assertTrue(copiedTable.behavior().insertRow());
      assertFalse(copiedTable.behavior().insertRowShift());
      assertTrue(
          workbook.xssfWorkbook().getSheet("Replica").validateSheetPassword("gridgrind-copy"));
    }
  }

  @Test
  void copySheetPreservesAdvancedCommentsAfterSaveAndReopen() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-comment-reopen-");

    List<CellCommentReport> sourceComments;
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet source = seedAdvancedCopySource(workbook);
      sourceComments =
          source
              .comments()
              .comments(
                  List.of("E2").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.CellAddress::new)
                      .toList());

      workbook.sheets().copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd());
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XssfWorkbook reopened =
        XssfWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals(
          sourceComments,
          reopened
              .sheet("Replica")
              .comments()
              .comments(
                  List.of("E2").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.CellAddress::new)
                      .toList()));
    }
  }

  @Test
  void copySheetRetargetsSupportedConditionalFormattingFamiliesAndNullSortState()
      throws IOException {
    AutofilterAdapter autofilterAdapter = new AutofilterAdapter();
    TableAdapter tableController = new TableAdapter();

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet source = workbook.getOrCreateSheet("Source");
      source.cells().setCell("A1", new CellValue.Text("Owner"));
      source.cells().setCell("B1", new CellValue.Text("Amount"));
      source.cells().setCell("C1", new CellValue.Text("Flag"));
      source.cells().setCell("A2", new CellValue.Text("Ada"));
      source.cells().setCell("A3", new CellValue.Text("Lin"));
      source.cells().setCell("B2", new CellValue.NumberValue(2));
      source.cells().setCell("B3", new CellValue.NumberValue(4));
      source.cells().setCell("E1", new CellValue.Text("Desk"));
      source.cells().setCell("F1", new CellValue.Text("Region"));
      source.cells().setCell("E2", new CellValue.Text("A1"));
      source.cells().setCell("E3", new CellValue.Text("B1"));
      source.cells().setCell("E4", new CellValue.Text("Totals"));
      source.cells().setCell("F2", new CellValue.Text("North"));
      source.cells().setCell("F3", new CellValue.Text("South"));
      source.cells().setCell("F4", new CellValue.Text("2"));

      source
          .metadata()
          .setAutofilter(
              "A1:B3",
              List.of(
                  new AutofilterFilterColumnSpec(
                      0L, true, new AutofilterFilterCriterionSpec.Values(List.of("Ada"), false))),
              Optional.empty());
      var rawValidations = source.xssfSheet().getCTWorksheet().addNewDataValidations();
      CTDataValidation rawValidation = rawValidations.addNewDataValidation();
      rawValidation.setSqref(List.of("C2:C3"));
      rawValidation.setFormula1("SUM(Source!$B$2:$B$3)");
      rawValidations.setCount(rawValidations.sizeOfDataValidationArray());
      source
          .metadata()
          .setConditionalFormatting(
              List.of("C2:C3").stream()
                  .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                  .toList(),
              new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
                  List.of(
                      new ConditionalFormattingRuleSpec.FormulaRule(
                          "SUM(Source!$B$2:$B$3)>0", true, Optional.empty()),
                      new ConditionalFormattingRuleSpec.CellValueRule(
                          ExcelComparisonOperator.GREATER_THAN,
                          "SUM(Source!$B$2:$B$3)",
                          Optional.empty(),
                          true,
                          Optional.empty()),
                      new ConditionalFormattingRuleSpec.ColorScaleRule(
                          false,
                          List.of(
                              new ConditionalFormattingThresholdSpec.Min(),
                              new ConditionalFormattingThresholdSpec.Max()),
                          List.of(
                              ColorSpec.rgb(new RgbHex("#112233")),
                              ColorSpec.rgb(new RgbHex("#445566")))),
                      new ConditionalFormattingRuleSpec.DataBarRule(
                          false,
                          ColorSpec.rgb(new RgbHex("#223344")),
                          false,
                          0,
                          100,
                          new ConditionalFormattingThresholdSpec.Min(),
                          new ConditionalFormattingThresholdSpec.Max()),
                      new ConditionalFormattingRuleSpec.IconSetRule(
                          false,
                          ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                          false,
                          true,
                          List.of(
                              new ConditionalFormattingThresholdSpec.Percent(0.0d),
                              new ConditionalFormattingThresholdSpec.Percent(50.0d),
                              new ConditionalFormattingThresholdSpec.Max())),
                      new ConditionalFormattingRuleSpec.Top10Rule(
                          true, 5, false, false, Optional.empty()))));
      new TableAdapter()
          .setTable(
              workbook,
              new ResolvedTable(
                  new TableName("Queue"),
                  new SheetName("Source"),
                  new RangeRef("E1:F4"),
                  true,
                  false,
                  new TableStyleSpec.None(),
                  "",
                  false,
                  false,
                  false,
                  "",
                  "",
                  "",
                  List.of()));

      workbook.sheets().copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd());

      assertEquals(
          List.of(
              new AutofilterEntryReport.SheetOwned(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                  List.of(
                      new AutofilterFilterColumnReport(
                          0L,
                          true,
                          new AutofilterFilterCriterionReport.Values(List.of("Ada"), false))),
                  java.util.Optional.empty())),
          autofilterAdapter.sheetOwnedAutofilters(workbook.sheet("Replica").xssfSheet()));
      assertEquals(
          "SUM(Replica!$B$2:$B$3)",
          workbook
              .sheet("Replica")
              .xssfSheet()
              .getCTWorksheet()
              .getDataValidations()
              .getDataValidationArray(0)
              .getFormula1());

      ConditionalFormattingEntryReport copiedFormatting =
          workbook.sheet("Replica").metadata().conditionalFormatting().getFirst();
      assertEquals(
          "SUM(Replica!$B$2:$B$3)>0",
          assertInstanceOf(
                  ConditionalFormattingRuleReport.FormulaRule.class,
                  copiedFormatting.rules().get(0))
              .formula());
      ConditionalFormattingRuleReport.CellValueRule copiedCellValueRule =
          assertInstanceOf(
              ConditionalFormattingRuleReport.CellValueRule.class, copiedFormatting.rules().get(1));
      assertEquals("SUM(Replica!$B$2:$B$3)", copiedCellValueRule.formula1());
      assertTrue(copiedCellValueRule.formula2().isEmpty());
      assertInstanceOf(
          ConditionalFormattingRuleReport.ColorScaleRule.class, copiedFormatting.rules().get(2));
      assertInstanceOf(
          ConditionalFormattingRuleReport.DataBarRule.class, copiedFormatting.rules().get(3));
      ConditionalFormattingRuleReport.IconSetRule copiedIconSetRule =
          assertInstanceOf(
              ConditionalFormattingRuleReport.IconSetRule.class, copiedFormatting.rules().get(4));
      assertTrue(copiedIconSetRule.reversed());
      assertInstanceOf(
          ConditionalFormattingRuleReport.Top10Rule.class, copiedFormatting.rules().get(5));

      TableEntryReport copiedTable =
          tableController.tables(workbook, new TableSelector.All()).stream()
              .filter(table -> "Replica".equals(table.sheetName().value()))
              .findFirst()
              .orElseThrow();
      assertEquals("E1:F4", copiedTable.range().value());
      assertEquals(1, copiedTable.structure().totalsRowCount());
      assertFalse(copiedTable.behavior().hasAutofilter());
      assertInstanceOf(TableStyleReport.None.class, copiedTable.style());
    }
  }

  @Test
  void localNamedRangeHelpersRespectScopeBoundaries() {
    NamedRangeReport.FormulaReport localFormula =
        new NamedRangeReport.FormulaReport(
            new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("LocalFormula"),
            new NamedRangeScopeReport.Sheet(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Source")),
            "SUM(Source!$A$1:$A$2)");
    List<NamedRangeReport> namedRanges =
        List.of(
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("WorkbookBudget"),
                new NamedRangeScopeReport.Workbook(),
                "Source!$A$1",
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Source"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1"))),
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("LocalBudget"),
                new NamedRangeScopeReport.Sheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Source")),
                "Source!$A$1:$A$2",
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Source"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:A2"))),
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("OtherLocal"),
                new NamedRangeScopeReport.Sheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Other")),
                "Other!$B$1",
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Other"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("B1"))),
            new NamedRangeReport.FormulaReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("OtherFormula"),
                new NamedRangeScopeReport.Sheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Other")),
                "SUM(Other!$A$1)"),
            new NamedRangeReport.FormulaReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("WorkbookFormula"),
                new NamedRangeScopeReport.Workbook(),
                "SUM(Source!$A$1)"),
            localFormula);

    assertEquals(
        List.of(
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("LocalBudget"),
                new NamedRangeScopeReport.Sheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Source")),
                "Source!$A$1:$A$2",
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Source"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:A2"))),
            localFormula),
        SheetCopyAdapter.copyableLocalNames(namedRanges, "Source"));
  }

  @Test
  void copySheetPreservesAndRetargetsRawDataValidations() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Source");
      workbook
          .sheet("Source")
          .metadata()
          .setDataValidation(
              "A1:A3",
              new ResolvedDataValidation(
                  new DataValidationRuleSpec.FormulaList(
                      ComparisonFormulas.normalizeFormula("Source!$D$1:$D$3", "formula")),
                  false,
                  false,
                  Optional.empty(),
                  Optional.empty()));
      addRawValidation(
          workbook.xssfWorkbook().getSheet("Source"), "B1:B3", STDataValidationType.LIST, "\"\"");
      addRawValidation(
          workbook.xssfWorkbook().getSheet("Source"), "C1:C3", STDataValidationType.LIST, null);

      workbook.sheets().copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd());

      List<DataValidationEntryReport> replicaValidations =
          workbook.sheet("Replica").metadata().dataValidations();
      assertEquals(3, replicaValidations.size());
      DataValidationEntryReport.Supported formulaList =
          assertInstanceOf(DataValidationEntryReport.Supported.class, replicaValidations.get(0));
      assertEquals(
          new DataValidationRuleSpec.FormulaList(
              ComparisonFormulas.normalizeFormula("Replica!$D$1:$D$3", "formula")),
          formulaList.validation().rule());
      DataValidationEntryReport.Supported emptyExplicitList =
          assertInstanceOf(DataValidationEntryReport.Supported.class, replicaValidations.get(1));
      assertEquals(
          new DataValidationRuleSpec.ExplicitList(List.of()),
          emptyExplicitList.validation().rule());
      assertEquals(
          new DataValidationEntryReport.Unsupported(
              List.of("C1:C3").stream()
                  .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                  .toList(),
              "MISSING_FORMULA",
              "List validation is missing both explicit values and formula1."),
          replicaValidations.get(2));
      List<CTDataValidation> rawReplicaValidations =
          List.of(
              workbook
                  .xssfWorkbook()
                  .getSheet("Replica")
                  .getCTWorksheet()
                  .getDataValidations()
                  .getDataValidationArray());
      assertEquals("Replica!$D$1:$D$3", rawReplicaValidations.get(0).getFormula1());
      assertEquals("\"\"", rawReplicaValidations.get(1).getFormula1());
      assertFalse(rawReplicaValidations.get(2).isSetFormula1());
    }
  }

  @Test
  void copySheetHandlesEdgeCasesInValidationFormulaRetargeting() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Source");
      var sheet = workbook.xssfWorkbook().getSheet("Source");
      // List validation: formula1.length() < 2 — covers isQuotedListLiteral length < 2 branch.
      addRawValidation(sheet, "A1", STDataValidationType.LIST, "1");
      // List validation: starts with '"' but does not end with '"' — covers startsWith=true,
      // endsWith=false branch in isQuotedListLiteral (valid formula, not a quoted literal).
      addRawValidation(sheet, "B1", STDataValidationType.LIST, "\"prefix\"&Source!A1");
      // Non-list validation: no formula1 set — covers isSetFormula1()=false branch.
      addRawValidation(sheet, "C1", STDataValidationType.WHOLE, null);
      // Non-list validation: formula1 is blank — covers isSetFormula1()=true, isBlank()=true
      // branch.
      addRawValidation(sheet, "D1", STDataValidationType.WHOLE, " ");
      // Non-list validation: formula1 non-blank, formula2 not set — covers isSetFormula2()=false.
      addRawValidation(sheet, "E1", STDataValidationType.WHOLE, "1");
      // Non-list validation: formula1 non-blank, formula2 blank — covers isSetFormula2()=true,
      // isBlank()=true branch.
      addRawValidationWithFormulas(sheet, "F1", STDataValidationType.WHOLE, "1", " ");
      // Raw validation without a declared type still needs formula retargeting.
      var validations = sheet.getCTWorksheet().getDataValidations();
      CTDataValidation missingType = validations.addNewDataValidation();
      missingType.setSqref(List.of("G1"));
      missingType.setFormula1("Source!$A$1");
      validations.setCount(validations.sizeOfDataValidationArray());

      workbook.sheets().copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd());

      List<CTDataValidation> rawReplica =
          List.of(
              workbook
                  .xssfWorkbook()
                  .getSheet("Replica")
                  .getCTWorksheet()
                  .getDataValidations()
                  .getDataValidationArray());
      assertEquals(7, rawReplica.size());
      assertEquals("1", rawReplica.get(0).getFormula1());
      assertTrue(rawReplica.get(1).getFormula1().contains("Replica"));
      assertFalse(rawReplica.get(2).isSetFormula1());
      assertEquals(" ", rawReplica.get(3).getFormula1());
      assertEquals("1", rawReplica.get(4).getFormula1());
      assertFalse(rawReplica.get(4).isSetFormula2());
      assertEquals("1", rawReplica.get(5).getFormula1());
      assertEquals(" ", rawReplica.get(5).getFormula2());
      assertEquals("Replica!$A$1", rawReplica.get(6).getFormula1());
    }
  }

  @Test
  void sheetOwnedAutofilterHandlesEmptySheetOwnedAndInvalidTableOwnedStates() {
    assertEquals(Optional.empty(), SheetCopyAdapter.sheetOwnedAutofilter(List.of()));
    assertEquals(
        Optional.of(
            new AutofilterEntryReport.SheetOwned(
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"))),
        SheetCopyAdapter.sheetOwnedAutofilter(
            List.of(
                new AutofilterEntryReport.SheetOwned(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3")))));
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                SheetCopyAdapter.sheetOwnedAutofilter(
                    List.of(
                        new AutofilterEntryReport.TableOwned(
                            new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"), "Ops"))));
    assertEquals(
        "sheetOwnedAutofilters must not return table-owned autofilter snapshots",
        failure.getMessage());
  }

  private static void addRawValidation(
      org.apache.poi.xssf.usermodel.XSSFSheet sheet,
      String range,
      STDataValidationType.Enum type,
      String formula1) {
    addRawValidationWithFormulas(sheet, range, type, formula1, null);
  }

  private static void addRawValidationWithFormulas(
      org.apache.poi.xssf.usermodel.XSSFSheet sheet,
      String range,
      STDataValidationType.Enum type,
      String formula1,
      String formula2) {
    var validations =
        sheet.getCTWorksheet().isSetDataValidations()
            ? sheet.getCTWorksheet().getDataValidations()
            : sheet.getCTWorksheet().addNewDataValidations();
    CTDataValidation validation = validations.addNewDataValidation();
    validation.setSqref(List.of(range));
    validation.setType(type);
    if (formula1 != null) {
      validation.setFormula1(formula1);
    }
    if (formula2 != null) {
      validation.setFormula2(formula2);
    }
    validations.setCount(validations.sizeOfDataValidationArray());
  }

  private static XssfSheet seedAdvancedCopySource(XssfWorkbook workbook) throws IOException {
    XssfSheet source = workbook.getOrCreateSheet("Source");
    source.cells().setCell("A1", new CellValue.Text("Owner"));
    source.cells().setCell("B1", new CellValue.Text("Amount"));
    source.cells().setCell("C1", new CellValue.Text("Formula"));
    source.cells().setCell("D1", new CellValue.Text("Stage"));
    source.cells().setCell("E1", new CellValue.Text("Comment"));
    source.cells().setCell("F1", new CellValue.Text("Flag"));
    source.cells().setCell("G1", new CellValue.Text("List"));
    source.cells().setCell("H1", new CellValue.Text("Whole"));
    source.cells().setCell("I1", new CellValue.Text("Note"));
    source.cells().setCell("J1", new CellValue.Text("Statuses"));
    source.cells().setCell("A2", new CellValue.Text("Ada"));
    source.cells().setCell("A3", new CellValue.Text("Lin"));
    source.cells().setCell("B2", new CellValue.NumberValue(2));
    source.cells().setCell("B3", new CellValue.NumberValue(4));
    source.cells().setCell("C2", new ResolvedFormula("SUM(Source!B2:B3)"));
    source.cells().setCell("D2", new CellValue.Text("Today"));
    source.cells().setCell("D3", new CellValue.Text("Today"));
    source.cells().setCell("F2", new CellValue.Text("High"));
    source.cells().setCell("F3", new CellValue.Text("Low"));
    source.cells().setCell("J2", new CellValue.Text("Ready"));
    source.cells().setCell("J3", new CellValue.Text("Done"));
    source.cells().setCell("H2", new CellValue.NumberValue(1));
    source.cells().setCell("H3", new CellValue.NumberValue(2));
    source.cells().setCell("E2", new CellValue.Text("Quarterly Review"));
    source
        .comments()
        .setComment(
            "E2",
            new ResolvedComment(
                "Quarterly Review",
                "GridGrind",
                true,
                Optional.of(
                    new ResolvedRichText(
                        List.of(
                            new ResolvedRichTextRun("Quarterly", Optional.empty()),
                            new ResolvedRichTextRun(
                                " Review",
                                Optional.of(
                                    new CellFontPatch(
                                        Optional.of(true),
                                        Optional.of(false),
                                        Optional.of("Aptos"),
                                        Optional.empty(),
                                        Optional.of(ColorSpec.rgb(new RgbHex("#112233"))),
                                        Optional.empty(),
                                        Optional.empty())))))),
                Optional.of(new CommentAnchorSpec(1, 1, 4, 5))));
    source
        .metadata()
        .setAutofilter(
            "A1:F3", advancedAutofilterCriteria(), Optional.of(advancedAutofilterSortState()));
    source
        .metadata()
        .setDataValidation(
            "G2:G3",
            new ResolvedDataValidation(
                new DataValidationRuleSpec.FormulaList(
                    ComparisonFormulas.normalizeFormula("Source!$J$2:$J$3", "formula")),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    source
        .metadata()
        .setDataValidation(
            "H2:H3",
            new ResolvedDataValidation(
                new DataValidationRuleSpec.WholeNumber(
                    ExcelComparisonOperator.BETWEEN,
                    ComparisonFormulas.normalizeFormula("1", "formula1"),
                    ComparisonFormulas.normalizeOptionalFormula(
                        Optional.of("SUM(Source!$B$2:$B$3)"), "formula2")),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    source
        .metadata()
        .setConditionalFormatting(
            List.of("I2:I3").stream()
                .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                .toList(),
            new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
                List.of(
                    new ConditionalFormattingRuleSpec.FormulaRule(
                        "SUM(Source!$B$2:$B$3)>0", false, Optional.of(supportedStyle())),
                    new ConditionalFormattingRuleSpec.CellValueRule(
                        ExcelComparisonOperator.BETWEEN,
                        "1",
                        Optional.of("SUM(Source!$B$2:$B$3)"),
                        false,
                        Optional.of(supportedStyle())))));
    workbook
        .names()
        .setNamedRange(
            new StructuredMutationAction.SetNamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName("LocalRange"),
                new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Source")),
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Source"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A2:A3"))));
    workbook
        .names()
        .setNamedRange(
            new StructuredMutationAction.SetNamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName("LocalFormula"),
                new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Source")),
                NamedRangeTarget.formula("SUM(Source!$B$2:$B$3)")));
    source.cells().setCell("L1", new CellValue.Text("Region"));
    source.cells().setCell("M1", new CellValue.Text("Desk"));
    source.cells().setCell("L2", new CellValue.Text("North"));
    source.cells().setCell("M2", new CellValue.Text("A1"));
    source.cells().setCell("L3", new CellValue.Text("South"));
    source.cells().setCell("M3", new CellValue.Text("B1"));
    new TableAdapter()
        .setTable(
            workbook,
            new ResolvedTable(
                new TableName("ReplicaTable"),
                new SheetName("Source"),
                new RangeRef("L1:M3"),
                false,
                true,
                new TableStyleSpec.Named("TableStyleMedium2", false, false, true, false),
                "replica table",
                true,
                true,
                false,
                "HeaderStyle",
                "DataStyle",
                "TotalsStyle",
                List.of(
                    new TableColumnSpec(0, "Region", "", "", ""),
                    new TableColumnSpec(1, "Desk", "", "", "UPPER([@Desk])"))));
    workbook
        .protection()
        .setSheetProtection("Source", protectionSettings(), Optional.of("gridgrind-copy"));
    return source;
  }

  private static List<AutofilterFilterColumnSpec> advancedAutofilterCriteria() {
    return List.of(
        new AutofilterFilterColumnSpec(
            0L, false, new AutofilterFilterCriterionSpec.Values(List.of("Ada"), true)),
        new AutofilterFilterColumnSpec(
            1L,
            true,
            new AutofilterFilterCriterionSpec.Custom(
                true,
                List.of(
                    new AutofilterFilterCriterionSpec.CustomConditionSpec("greaterThan", "1"),
                    new AutofilterFilterCriterionSpec.CustomConditionSpec("equal", "Queue")))),
        new AutofilterFilterColumnSpec(
            2L,
            true,
            new AutofilterFilterCriterionSpec.Dynamic(
                "today", Optional.of(1.0d), Optional.of(2.0d))),
        new AutofilterFilterColumnSpec(
            3L, true, new AutofilterFilterCriterionSpec.Top10(10, false, true)),
        new AutofilterFilterColumnSpec(
            4L,
            true,
            new AutofilterFilterCriterionSpec.Color(false, ColorSpec.rgb(new RgbHex("#AABBCC")))),
        new AutofilterFilterColumnSpec(
            5L, true, new AutofilterFilterCriterionSpec.Icon("3TrafficLights1", 2)));
  }

  private static AutofilterSortStateSpec advancedAutofilterSortState() {
    return new AutofilterSortStateSpec(
        new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:F3"),
        true,
        true,
        java.util.Optional.of(dev.erst.gridgrind.core.vocabulary.ExcelAutofilterSortMethod.STROKE),
        List.of(
            new AutofilterSortConditionSpec.CellColor(
                new dev.erst.gridgrind.core.vocabulary.RangeRef("A2:A3"),
                true,
                ColorSpec.rgb(new RgbHex("#102030"))),
            new AutofilterSortConditionSpec.Icon(
                new dev.erst.gridgrind.core.vocabulary.RangeRef("B2:B3"), false, 4)));
  }

  private static DifferentialStylePatch supportedStyle() {
    return new DifferentialStylePatch(
        Optional.of("0.00"),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private static SheetProtectionSettings protectionSettings() {
    return new SheetProtectionSettings(
        true, false, true, false, true, false, true, false, true, false, true, false, true, false,
        true);
  }
}
