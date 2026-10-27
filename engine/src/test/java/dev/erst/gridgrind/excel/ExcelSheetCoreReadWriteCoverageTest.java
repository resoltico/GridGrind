package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.layout.PaneSpec;
import dev.erst.gridgrind.core.model.layout.SheetDefaultsSpec;
import dev.erst.gridgrind.core.model.layout.SheetDisplaySpec;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummarySpec;
import dev.erst.gridgrind.core.model.layout.SheetPresentationSpec;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorSpec;
import dev.erst.gridgrind.core.model.value.CellHyperlinkReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.plan.CellReadFacet;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion;
import dev.erst.gridgrind.core.vocabulary.ExcelRowSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.PaneType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** ExcelSheet core read, write, parsing, and facade coverage. */
class ExcelSheetCoreReadWriteCoverageTest extends ExcelSheetTestSupport {
  @Test
  void readsWritesAndSnapshotsTypedCellsAndFormulaResults() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);
      ExcelSheetCells cells = sheet.cells();

      assertSame(
          cells,
          cells.appendRow(
              new CellValue.Text("Name"),
              new CellValue.NumberValue(42.5),
              new CellValue.BooleanValue(true),
              new ResolvedFormula("B1*2"),
              new ResolvedFormula("TRUE()"),
              new ResolvedFormula("\"Hi\"")));
      sheet
          .cells()
          .appendRow(
              new CellValue.Blank(),
              new CellValueSpec.Date(LocalDate.of(2026, 3, 23)),
              new CellValueSpec.DateTime(LocalDateTime.of(2026, 3, 23, 14, 15, 16)),
              new ResolvedFormula("1/0"));
      sheet.cells().setCell("A3", new ResolvedFormula("B3+1"));
      sheet.cells().setCell("B3", new CellValueSpec.Date(LocalDate.of(2026, 3, 24)));
      sheet.cells().setCell("C3", new CellValueSpec.DateTime(LocalDateTime.of(2026, 3, 24, 9, 0)));
      sheet.cells().setCell("D3", new ResolvedFormula("D3+1"));
      sheet.columns().autoSize();

      Row errorRow = poiSheet.createRow(3);
      Cell errorCell = errorRow.createCell(0);
      errorCell.setCellErrorValue(FormulaError.DIV0.getCode());

      assertEquals("Budget", sheet.name());
      assertEquals("Name", sheet.cells().text("A1"));
      assertEquals(85.0, sheet.cells().number("D1"));
      assertTrue(sheet.cells().bool("C1"));
      assertTrue(sheet.cells().bool("E1"));
      assertEquals("B1*2", sheet.cells().formula("D1"));
      assertEquals(4, sheet.rows().physicalCount());
      assertEquals(3, sheet.rows().lastIndex());
      assertEquals(5, sheet.columns().lastIndex());

      CellReport.TextReport textSnapshot = (CellReport.TextReport) sheet.cells().snapshotCell("A1");
      assertEquals("TEXT", textSnapshot.type());
      assertEquals("Name", textSnapshot.value().orElseThrow().value());
      assertTrue(textSnapshot.runs().isEmpty());

      CellReport.NumberReport numberSnapshot =
          (CellReport.NumberReport) sheet.cells().snapshotCell("B1");
      assertEquals("NUMBER", numberSnapshot.type());
      assertEquals(42.5, numberSnapshot.value().orElseThrow().value());

      CellReport.BooleanReport booleanSnapshot =
          (CellReport.BooleanReport) sheet.cells().snapshotCell("C1");
      assertEquals("BOOLEAN", booleanSnapshot.type());
      assertTrue(booleanSnapshot.value().orElseThrow().value());

      CellReport.BlankReport blankSnapshot =
          (CellReport.BlankReport) sheet.cells().snapshotCell("A2");
      assertEquals("BLANK", blankSnapshot.type());

      CellReport.FormulaReport stringFormulaSnapshot =
          (CellReport.FormulaReport) sheet.cells().snapshotCell("F1");
      assertEquals("FORMULA", stringFormulaSnapshot.type());
      assertEquals("\"Hi\"", stringFormulaSnapshot.formula().orElseThrow());
      assertEquals(
          "Hi",
          ((dev.erst.gridgrind.core.model.value.CellValue.Text)
                  stringFormulaSnapshot.value().orElseThrow())
              .value());

      CellReport.FormulaReport errorFormulaSnapshot =
          (CellReport.FormulaReport) sheet.cells().snapshotCell("D2");
      assertEquals("FORMULA", errorFormulaSnapshot.type());
      assertEquals(
          "#DIV/0!",
          ((dev.erst.gridgrind.core.model.value.CellValue.ErrorValue)
                  errorFormulaSnapshot.value().orElseThrow())
              .value());

      CellReport.ErrorReport errorSnapshot =
          (CellReport.ErrorReport) sheet.cells().snapshotCell("A4");
      assertEquals("ERROR", errorSnapshot.type());
      assertEquals("#DIV/0!", errorSnapshot.value().orElseThrow().value());
      CellReport.FormulaReport circularFormulaSnapshot =
          (CellReport.FormulaReport) sheet.cells().snapshotCell("D3");
      assertEquals(
          "#CIRCULAR_REF!",
          ((dev.erst.gridgrind.core.model.value.CellValue.ErrorValue)
                  circularFormulaSnapshot.value().orElseThrow())
              .value());

      List<ExcelPreviewRow> preview = sheet.cells().preview(4, 6);
      assertEquals(4, preview.size());
      assertEquals("A1", preview.get(0).cells().get(0).address().value());
      assertTrue(
          preview.get(1).cells().stream().noneMatch(cell -> "A2".equals(cell.address().value())));
      assertEquals("Hi", preview.get(0).cells().get(5).displayValue().orElseThrow());
    }
  }

  @Test
  void projectionAwareSnapshotReadsSkipOrIncludeFormulaEvaluation() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheetCells cells =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator).cells();
      cells.setCell("A1", new ResolvedFormula("1+1"));

      CellReadProjection formulaOnly =
          new CellReadProjection(List.copyOf(Set.of(CellReadFacet.FORMULA)));
      CellReadProjection formatOnly =
          new CellReadProjection(List.copyOf(Set.of(CellReadFacet.FORMAT)));
      CellReport.FormulaReport formulaOnlySnapshot =
          (CellReport.FormulaReport) cells.snapshotCell("A1", formulaOnly);
      CellReport.FormulaReport formatOnlySnapshot =
          (CellReport.FormulaReport) cells.snapshotCell("A1", formatOnly);

      assertTrue(formulaOnlySnapshot.value().isEmpty());
      assertTrue(formatOnlySnapshot.displayValue().isPresent());
      assertTrue(formatOnlySnapshot.value().isEmpty());
      assertEquals(1, cells.snapshotCells(List.of("A1")).size());
    }
  }

  @Test
  void validatesWriteOperationArguments() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      assertThrows(
          NullPointerException.class, () -> sheet.cells().setCell(null, new CellValue.Text("x")));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.cells().setCell(" ", new CellValue.Text("x")));
      assertThrows(NullPointerException.class, () -> sheet.cells().setCell("A1", null));
      assertThrows(
          NullPointerException.class,
          () -> sheet.cells().setRange(null, List.of(List.of(new CellValue.Text("x")))));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.cells().setRange(" ", List.of(List.of(new CellValue.Text("x")))));
      assertThrows(NullPointerException.class, () -> sheet.cells().setRange("A1", null));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().setRange("A1", List.of()));
      assertThrows(NullPointerException.class, () -> sheet.cells().clearRange(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().clearRange(" "));
      assertThrows(NullPointerException.class, () -> sheet.layout().mergeCells(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.layout().mergeCells(" "));
      assertThrows(NullPointerException.class, () -> sheet.layout().unmergeCells(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.layout().unmergeCells(" "));
      assertThrows(IllegalArgumentException.class, () -> sheet.columns().setWidth(-1, 0, 16.0));
      assertThrows(IllegalArgumentException.class, () -> sheet.columns().setWidth(1, 0, 16.0));
      assertThrows(IllegalArgumentException.class, () -> sheet.columns().setWidth(0, 0, 0.0));
      assertThrows(IllegalArgumentException.class, () -> sheet.columns().setWidth(0, 0, 256.0));
      assertThrows(
          IllegalArgumentException.class, () -> sheet.columns().setWidth(0, 0, Double.MIN_VALUE));
      assertThrows(
          IllegalArgumentException.class, () -> sheet.columns().setWidth(0, 0, Double.NaN));
      assertThrows(IllegalArgumentException.class, () -> sheet.rows().setHeight(-1, 0, 28.5));
      assertThrows(IllegalArgumentException.class, () -> sheet.rows().setHeight(1, 0, 28.5));
      assertThrows(IllegalArgumentException.class, () -> sheet.rows().setHeight(0, 0, 0.0));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet
                  .rows()
                  .setHeight(0, 0, Math.nextUp(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS)));
      assertThrows(
          IllegalArgumentException.class, () -> sheet.rows().setHeight(0, 0, Double.MIN_VALUE));
      assertThrows(IllegalArgumentException.class, () -> sheet.rows().setHeight(0, 0, Double.NaN));
      assertThrows(NullPointerException.class, () -> sheet.layout().setPane(null));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.layout().setPane(new PaneSpec.Frozen(-1, 0, 0, 0)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.layout().setPane(new PaneSpec.Frozen(0, 0, 0, 0)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.layout().setPane(new PaneSpec.Frozen(0, 1, 1, 1)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.layout().setPane(new PaneSpec.Frozen(1, 0, 1, 1)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.layout().setPane(new PaneSpec.Frozen(2, 1, 1, 1)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.layout().setPane(new PaneSpec.Frozen(1, 2, 1, 1)));
      assertThrows(IllegalArgumentException.class, () -> sheet.layout().setZoom(9));
      assertThrows(IllegalArgumentException.class, () -> sheet.layout().setZoom(401));
      assertThrows(NullPointerException.class, () -> sheet.layout().setPrintLayout(null));
      assertThrows(
          NullPointerException.class,
          () ->
              sheet
                  .cells()
                  .applyStyle(
                      null,
                      new CellStylePatch(
                          Optional.of("0"),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty())));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet
                  .cells()
                  .applyStyle(
                      " ",
                      new CellStylePatch(
                          Optional.of("0"),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty())));
      assertThrows(NullPointerException.class, () -> sheet.cells().applyStyle("A1", null));
      assertThrows(
          NullPointerException.class,
          () ->
              sheet
                  .annotations()
                  .setHyperlink(null, new HyperlinkTarget.Url("https://example.com")));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet
                  .annotations()
                  .setHyperlink(" ", new HyperlinkTarget.Url("https://example.com")));
      assertThrows(NullPointerException.class, () -> sheet.annotations().setHyperlink("A1", null));
      assertThrows(NullPointerException.class, () -> sheet.annotations().clearHyperlink(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.annotations().clearHyperlink(" "));
      assertThrows(
          NullPointerException.class,
          () ->
              sheet
                  .annotations()
                  .setComment(null, new ResolvedComment("Review", "GridGrind", false)));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet
                  .annotations()
                  .setComment(" ", new ResolvedComment("Review", "GridGrind", false)));
      assertThrows(NullPointerException.class, () -> sheet.annotations().setComment("A1", null));
      assertThrows(NullPointerException.class, () -> sheet.annotations().clearComment(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.annotations().clearComment(" "));
      assertThrows(
          NullPointerException.class, () -> sheet.cells().appendRow((ResolvedCellWrite[]) null));
      assertThrows(
          NullPointerException.class, () -> sheet.cells().appendRow(new CellValue.Text("x"), null));
    }
  }

  @Test
  void validatesAddressAndRangeParsing() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      assertThrows(NullPointerException.class, () -> sheet.cells().snapshotCell(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().snapshotCell(" "));
      InvalidCellAddressException invalidSetCell =
          assertThrows(
              InvalidCellAddressException.class,
              () -> sheet.cells().setCell(":", new CellValue.Text("x")));
      assertEquals(":", invalidSetCell.address());
      InvalidCellAddressException invalidSnapshotCell =
          assertThrows(InvalidCellAddressException.class, () -> sheet.cells().snapshotCell(":"));
      assertEquals(":", invalidSnapshotCell.address());
      InvalidCellAddressException badAddrSnapshot =
          assertThrows(
              InvalidCellAddressException.class, () -> sheet.cells().snapshotCell("BADADDR"));
      assertEquals("BADADDR", badAddrSnapshot.address());
      InvalidCellAddressException a0Snapshot =
          assertThrows(InvalidCellAddressException.class, () -> sheet.cells().snapshotCell("A0"));
      assertEquals("A0", a0Snapshot.address());
      InvalidCellAddressException numericOnlySnapshot =
          assertThrows(InvalidCellAddressException.class, () -> sheet.cells().snapshotCell("1"));
      assertEquals("1", numericOnlySnapshot.address());
      InvalidCellAddressException outOfBoundsRow =
          assertThrows(
              InvalidCellAddressException.class, () -> sheet.cells().snapshotCell("A1048577"));
      assertEquals("A1048577", outOfBoundsRow.address());
      InvalidCellAddressException outOfBoundsCol =
          assertThrows(InvalidCellAddressException.class, () -> sheet.cells().snapshotCell("XFE1"));
      assertEquals("XFE1", outOfBoundsCol.address());
      InvalidRangeAddressException invalidRangeSet =
          assertThrows(
              InvalidRangeAddressException.class,
              () -> sheet.cells().setRange("A1:", List.of(List.of(new CellValue.Text("x")))));
      assertEquals("A1:", invalidRangeSet.range());
      InvalidRangeAddressException invalidRangeClear =
          assertThrows(
              InvalidRangeAddressException.class, () -> sheet.cells().clearRange("A1:B2:C3"));
      assertEquals("A1:B2:C3", invalidRangeClear.range());
      InvalidRangeAddressException invalidRangeMerge =
          assertThrows(InvalidRangeAddressException.class, () -> sheet.layout().mergeCells("A1:"));
      assertEquals("A1:", invalidRangeMerge.range());
      InvalidRangeAddressException invalidRangeUnmerge =
          assertThrows(
              InvalidRangeAddressException.class, () -> sheet.layout().unmergeCells("A1:"));
      assertEquals("A1:", invalidRangeUnmerge.range());
      assertThrows(
          InvalidRangeAddressException.class,
          () ->
              sheet
                  .cells()
                  .applyStyle(
                      "A1:",
                      new CellStylePatch(
                          Optional.of("0"),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty())));
      assertThrows(IllegalArgumentException.class, () -> sheet.layout().mergeCells("A1"));
      assertThrows(IllegalArgumentException.class, () -> sheet.layout().unmergeCells("A1:B2"));
    }
  }

  @Test
  void groupsAndUngroupsRowsAndColumnsThroughSheetFacade() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);
      ExcelSheetRows rows = sheet.rows();
      ExcelSheetColumns columns = sheet.columns();

      sheet.cells().setCell("A1", new CellValue.Text("Header"));

      assertSame(rows, rows.group(new ExcelRowSpan(1, 3), true));
      assertSame(columns, columns.group(new ExcelColumnSpan(1, 3), true));
      assertSame(rows, rows.ungroup(new ExcelRowSpan(1, 3)));
      assertSame(columns, columns.ungroup(new ExcelColumnSpan(1, 3)));

      SheetLayoutReport layout = sheet.layout().snapshot();
      assertFalse(layout.rows().get(1).hidden());
      assertEquals(0, layout.rows().get(1).outlineLevel());
      assertFalse(layout.columns().get(1).hidden());
      assertEquals(0, layout.columns().get(1).outlineLevel());
    }
  }

  @Test
  void validatesPreviewAndReadFailures() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      assertThrows(IllegalArgumentException.class, () -> sheet.cells().preview(0, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().preview(1, 0));
      assertEquals(List.of(), sheet.cells().preview(3, 3));
      CellNotFoundException missingCell =
          assertThrows(CellNotFoundException.class, () -> sheet.cells().text("A1"));
      assertEquals("A1", missingCell.address());
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().text(" "));

      sheet.cells().setCell("A1", new CellValue.Text("Name"));
      sheet.cells().setCell("B1", new ResolvedFormula("TRUE()"));
      sheet.cells().setCell("C1", new ResolvedFormula("1+1"));

      assertThrows(CellNotFoundException.class, () -> sheet.cells().text("B2"));
      assertThrows(CellNotFoundException.class, () -> sheet.cells().text("D1"));
      assertThrows(IllegalStateException.class, () -> sheet.cells().number("A1"));
      assertThrows(IllegalStateException.class, () -> sheet.cells().number("B1"));
      assertThrows(IllegalStateException.class, () -> sheet.cells().bool("C1"));
      assertThrows(IllegalStateException.class, () -> sheet.cells().formula("A1"));
    }
  }

  @Test
  void supportsRangeWritesClearsStylesAndStyleAwarePreview() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);
      ExcelSheetCells cells = sheet.cells();

      assertSame(
          cells,
          cells.setRange(
              "B2:A1",
              List.of(
                  List.of(new CellValue.Text("Item"), new CellValue.NumberValue(42.0)),
                  List.of(new CellValue.Text("Tax"), new CellValue.NumberValue(8.0)))));
      sheet
          .cells()
          .applyStyle(
              "A1:B1",
              new CellStylePatch(
                  Optional.of("#,##0.00"),
                  Optional.of(
                      new CellAlignmentPatch(
                          Optional.of(true),
                          Optional.of(ExcelHorizontalAlignment.CENTER),
                          Optional.of(ExcelVerticalAlignment.TOP),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.of(
                      new CellFontPatch(
                          Optional.of(true),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.empty(),
                  Optional.empty(),
                  Optional.empty()));
      sheet
          .cells()
          .applyStyle(
              "C1",
              new CellStylePatch(
                  Optional.empty(),
                  Optional.empty(),
                  Optional.of(
                      new CellFontPatch(
                          Optional.empty(),
                          Optional.of(true),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.empty(),
                  Optional.empty(),
                  Optional.empty()));

      CellReport styledValue = sheet.cells().snapshotCell("A1");
      assertEquals("#,##0.00", styledValue.style().orElseThrow().numberFormat());
      assertTrue(styledValue.style().orElseThrow().font().bold());
      assertTrue(styledValue.style().orElseThrow().alignment().wrapText());
      assertEquals(
          ExcelHorizontalAlignment.CENTER,
          styledValue.style().orElseThrow().alignment().horizontalAlignment());
      assertEquals(
          ExcelVerticalAlignment.TOP,
          styledValue.style().orElseThrow().alignment().verticalAlignment());
      assertEquals("Calibri", styledValue.style().orElseThrow().font().fontName());
      assertEquals(220, styledValue.style().orElseThrow().font().fontHeight().twips());
      assertFalse(styledValue.style().orElseThrow().font().underline());
      assertFalse(styledValue.style().orElseThrow().font().strikeout());
      assertEquals(
          CellColorReport.indexed(8), styledValue.style().orElseThrow().font().fontColor());
      assertNull(fillForegroundColor(styledValue.style().orElseThrow().fill()));
      assertEquals(
          ExcelBorderStyle.NONE, borderStyle(styledValue.style().orElseThrow().border().top()));
      assertEquals(
          ExcelBorderStyle.NONE, borderStyle(styledValue.style().orElseThrow().border().right()));
      assertEquals(
          ExcelBorderStyle.NONE, borderStyle(styledValue.style().orElseThrow().border().bottom()));
      assertEquals(
          ExcelBorderStyle.NONE, borderStyle(styledValue.style().orElseThrow().border().left()));

      List<ExcelPreviewRow> preview = sheet.cells().preview(2, 3);
      assertTrue(
          preview.getFirst().cells().stream()
              .anyMatch(cell -> "C1".equals(cell.address().value())));
      assertEquals("BLANK", sheet.cells().snapshotCell("C1").type());
      assertTrue(sheet.cells().snapshotCell("C1").style().orElseThrow().font().italic());

      sheet.cells().clearRange("A2:B2");

      CellReport cleared = sheet.cells().snapshotCell("A2");
      assertEquals("BLANK", cleared.type());
      assertEquals("General", cleared.style().orElseThrow().numberFormat());
      assertFalse(cleared.style().orElseThrow().font().bold());
      assertEquals(
          ExcelHorizontalAlignment.GENERAL,
          cleared.style().orElseThrow().alignment().horizontalAlignment());
      assertEquals(
          ExcelVerticalAlignment.BOTTOM,
          cleared.style().orElseThrow().alignment().verticalAlignment());
      assertEquals("Calibri", cleared.style().orElseThrow().font().fontName());
      assertEquals(220, cleared.style().orElseThrow().font().fontHeight().twips());
      assertEquals(CellColorReport.indexed(8), cleared.style().orElseThrow().font().fontColor());
      assertFalse(cleared.style().orElseThrow().font().underline());
      assertFalse(cleared.style().orElseThrow().font().strikeout());
      assertNull(fillForegroundColor(cleared.style().orElseThrow().fill()));
      assertEquals(
          ExcelBorderStyle.NONE, borderStyle(cleared.style().orElseThrow().border().top()));
      assertEquals(
          ExcelBorderStyle.NONE, borderStyle(cleared.style().orElseThrow().border().right()));
      assertEquals(
          ExcelBorderStyle.NONE, borderStyle(cleared.style().orElseThrow().border().bottom()));
      assertEquals(
          ExcelBorderStyle.NONE, borderStyle(cleared.style().orElseThrow().border().left()));

      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.cells().setRange("A1:B2", List.of(List.of(new CellValue.Text("x")))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet
                  .cells()
                  .setRange("A1:B2", List.of(List.of(), List.of(new CellValue.Text("x")))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet
                  .cells()
                  .setRange(
                      "A1:B2",
                      List.of(
                          List.of(new CellValue.Text("x")),
                          List.of(new CellValue.Text("y"), new CellValue.Text("z")))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet
                  .cells()
                  .setRange(
                      "A1:B2",
                      List.of(List.of(new CellValue.Text("x")), List.of(new CellValue.Text("y")))));
      List<List<ResolvedCellWrite>> rowsWithNull = new ArrayList<>();
      List<ResolvedCellWrite> rowWithNull = new ArrayList<>();
      rowWithNull.add(null);
      rowsWithNull.add(rowWithNull);
      assertThrows(NullPointerException.class, () -> sheet.cells().setRange("A1", rowsWithNull));
    }
  }

  @Test
  void mergesFormattingDepthStylesAndPreservesExistingAttributes() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.cells().setCell("A1", new CellValue.Text("Item"));
      sheet
          .cells()
          .applyStyle(
              "A1",
              new CellStylePatch(
                  Optional.empty(),
                  Optional.of(
                      new CellAlignmentPatch(
                          Optional.of(true),
                          Optional.of(ExcelHorizontalAlignment.CENTER),
                          Optional.of(ExcelVerticalAlignment.TOP),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.of(
                      new CellFontPatch(
                          Optional.of(true),
                          Optional.of(false),
                          Optional.of("Aptos"),
                          Optional.of(new FontHeightSpec.Twips(280)),
                          Optional.of(ColorSpec.rgb(new RgbHex("#1F4E78"))),
                          Optional.of(true),
                          Optional.of(false))),
                  Optional.of(
                      CellFillSpec.patternForeground(
                          ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#FFF2CC")))),
                  Optional.of(
                      new CellBorderPatch(
                          Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.THIN)),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.empty()));
      sheet
          .cells()
          .applyStyle(
              "A1",
              new CellStylePatch(
                  Optional.empty(),
                  Optional.empty(),
                  Optional.of(
                      new CellFontPatch(
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.of(true))),
                  Optional.empty(),
                  Optional.of(
                      new CellBorderPatch(
                          Optional.empty(),
                          Optional.empty(),
                          Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.DOUBLE)),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.empty()));

      CellReport styled = sheet.cells().snapshotCell("A1");
      assertTrue(styled.style().orElseThrow().font().bold());
      assertFalse(styled.style().orElseThrow().font().italic());
      assertTrue(styled.style().orElseThrow().alignment().wrapText());
      assertEquals(
          ExcelHorizontalAlignment.CENTER,
          styled.style().orElseThrow().alignment().horizontalAlignment());
      assertEquals(
          ExcelVerticalAlignment.TOP, styled.style().orElseThrow().alignment().verticalAlignment());
      assertEquals("Aptos", styled.style().orElseThrow().font().fontName());
      assertEquals(280, styled.style().orElseThrow().font().fontHeight().twips());
      assertEquals(rgb("#1F4E78"), styled.style().orElseThrow().font().fontColor());
      assertTrue(styled.style().orElseThrow().font().underline());
      assertTrue(styled.style().orElseThrow().font().strikeout());
      assertEquals(rgb("#FFF2CC"), fillForegroundColor(styled.style().orElseThrow().fill()));
      assertEquals(ExcelBorderStyle.THIN, borderStyle(styled.style().orElseThrow().border().top()));
      assertEquals(
          ExcelBorderStyle.DOUBLE, borderStyle(styled.style().orElseThrow().border().right()));
      assertEquals(
          ExcelBorderStyle.THIN, borderStyle(styled.style().orElseThrow().border().bottom()));
      assertEquals(
          ExcelBorderStyle.THIN, borderStyle(styled.style().orElseThrow().border().left()));
    }
  }

  @Test
  void snapshotsAndClearsHyperlinksAndComments() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.annotations().setHyperlink("A1", new HyperlinkTarget.Url("https://example.com/report"));
      sheet.annotations().setComment("A1", new ResolvedComment("Review", "GridGrind", true));

      CellReport.BlankReport snapshot = (CellReport.BlankReport) sheet.cells().snapshotCell("A1");
      assertEquals(
          new HyperlinkTarget.Url("https://example.com/report"),
          snapshot.hyperlink().orElseThrow());
      assertAll(
          () -> assertEquals("Review", snapshot.comment().orElseThrow().text()),
          () -> assertEquals("GridGrind", snapshot.comment().orElseThrow().author()),
          () -> assertTrue(snapshot.comment().orElseThrow().visible()));

      List<ExcelPreviewRow> preview = sheet.cells().preview(1, 1);
      assertEquals(1, preview.size());
      assertEquals("A1", preview.getFirst().cells().getFirst().address().value());

      sheet.annotations().clearHyperlink("A1");
      sheet.annotations().clearComment("A1");
      CellReport.BlankReport clearedMetadata =
          (CellReport.BlankReport) sheet.cells().snapshotCell("A1");
      assertTrue(clearedMetadata.hyperlink().isEmpty());
      assertTrue(clearedMetadata.comment().isEmpty());

      sheet.annotations().setHyperlink("A1", new HyperlinkTarget.Document("Budget!B4"));
      sheet.annotations().setComment("A1", new ResolvedComment("Again", "GridGrind", false));
      sheet.cells().clearRange("A1");
      CellReport.BlankReport clearedRange =
          (CellReport.BlankReport) sheet.cells().snapshotCell("A1");
      assertTrue(clearedRange.hyperlink().isEmpty());
      assertTrue(clearedRange.comment().isEmpty());

      // clearHyperlink and clearComment are no-ops on cells that do not physically exist
      assertDoesNotThrow(() -> sheet.annotations().clearHyperlink("B2"));
      assertDoesNotThrow(() -> sheet.annotations().clearComment("B2"));
      // calling again on B2 (still non-existent) must still be a no-op, not throw
      assertDoesNotThrow(() -> sheet.annotations().clearHyperlink("B2"));
      assertDoesNotThrow(() -> sheet.annotations().clearComment("B2"));
    }
  }

  @Test
  void writesAndReadsFileHyperlinksWithSpacesInPlainPaths() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      assertDoesNotThrow(
          () ->
              sheet
                  .annotations()
                  .setHyperlink("A1", new HyperlinkTarget.File("support/budget backup.xlsx")));

      Cell cell = poiSheet.getRow(0).getCell(0);
      assertEquals("support/budget%20backup.xlsx", cell.getHyperlink().getAddress());
      assertEquals(
          Optional.of(new HyperlinkTarget.File("support/budget backup.xlsx")),
          ExcelSheetAnnotationSupport.hyperlink(cell));
    }
  }

  @Test
  void clearRangeIsNoOpOnNeverWrittenCells() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Data");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.cells().setCell("A1", new CellValue.Text("anchor"));

      int rowsBefore = sheet.rows().physicalCount();
      int lastRowBefore = sheet.rows().lastIndex();
      int lastColBefore = sheet.columns().lastIndex();

      // clear a range that has never been written
      sheet.cells().clearRange("B2:E5");

      // physicalRowCount, lastRowIndex, and lastColumnIndex must not change
      assertEquals(rowsBefore, sheet.rows().physicalCount(), "physicalRowCount must not increase");
      assertEquals(lastRowBefore, sheet.rows().lastIndex(), "lastRowIndex must not change");
      assertEquals(lastColBefore, sheet.columns().lastIndex(), "lastColumnIndex must not change");
    }
  }

  @Test
  void clearRangeSkipsAbsentCellsInExistingRows() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Data");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      // Row 2 exists (B2 is written) but C2 has never been written.
      // Clearing B2:C2 must not throw even though C2 is absent.
      sheet.cells().setCell("B2", new CellValue.Text("present"));
      assertDoesNotThrow(() -> sheet.cells().clearRange("B2:C2"));
      // B2 must now be blank after the clear
      CellReport b2 = sheet.cells().snapshotCell("B2");
      assertEquals("BLANK", b2.type());
    }
  }

  @Test
  void introspectsWindowsSelectionsAndLayoutAcrossSparseMetadata() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      SheetLayoutReport emptyLayout = sheet.layout().snapshot();
      assertEquals(new dev.erst.gridgrind.core.model.layout.PaneReport.None(), emptyLayout.pane());
      assertEquals(100, emptyLayout.zoomPercent());
      assertEquals(
          dev.erst.gridgrind.core.model.layout.SheetDisplayReport.defaults(),
          emptyLayout.presentation().display());
      assertEquals(
          dev.erst.gridgrind.core.model.layout.SheetOutlineSummaryReport.defaults(),
          emptyLayout.presentation().outlineSummary());
      assertEquals(
          dev.erst.gridgrind.core.model.layout.SheetDefaultsReport.defaults(),
          emptyLayout.presentation().sheetDefaults());
      assertEquals(List.of(), emptyLayout.columns());
      assertEquals(List.of(), emptyLayout.rows());

      sheet.cells().setCell("B2", new CellValue.Text("Center"));
      sheet.annotations().setHyperlink("A1", new HyperlinkTarget.Url("https://example.com/report"));
      sheet.annotations().setComment("C3", new ResolvedComment("Review", "GridGrind", false));
      sheet.columns().setWidth(0, 0, 12.5);
      sheet.rows().setHeight(0, 0, 19.5);
      sheet
          .layout()
          .setPresentation(
              new SheetPresentationSpec(
                  new SheetDisplaySpec(false, false, false, true, true),
                  Optional.of(ColorSpec.rgb(new RgbHex("#112233"))),
                  new SheetOutlineSummarySpec(false, false),
                  new SheetDefaultsSpec(11, 18.5d),
                  List.of(
                      new IgnoredErrorSpec(
                          new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B2"),
                          List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT)))));

      dev.erst.gridgrind.core.model.layout.WindowReport.Dense window =
          sheet.cells().window("A1", 3, 3);
      assertEquals("A1", window.rows().getFirst().cells().getFirst().address().value());
      assertEquals("B2", window.rows().get(1).cells().get(1).address().value());
      assertEquals("Center", window.rows().get(1).cells().get(1).displayValue().orElseThrow());
      assertEquals("C3", window.rows().get(2).cells().get(2).address().value());
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().window("A1", 0, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().window("A1", 1, 0));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().window("A1", 1048577, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().window("A1", 1, 16385));

      List<CellHyperlinkReport> allHyperlinks = sheet.annotations().hyperlinks();
      assertEquals(1, allHyperlinks.size());
      assertEquals("A1", allHyperlinks.getFirst().address().value());

      List<CellHyperlinkReport> selectedHyperlinks =
          sheet
              .annotations()
              .hyperlinks(
                  List.of("A1", "B2", "B9", "D3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.CellAddress::new)
                      .toList());
      assertEquals(1, selectedHyperlinks.size());
      assertEquals("A1", selectedHyperlinks.getFirst().address().value());

      List<CellCommentReport> selectedComments =
          sheet
              .annotations()
              .comments(
                  List.of("C3", "B2", "A9", "D3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.CellAddress::new)
                      .toList());
      assertEquals(1, selectedComments.size());
      assertEquals("C3", selectedComments.getFirst().address().value());

      SheetLayoutReport unfrozenLayout = sheet.layout().snapshot();
      assertEquals(
          new dev.erst.gridgrind.core.model.layout.PaneReport.None(), unfrozenLayout.pane());
      assertEquals(100, unfrozenLayout.zoomPercent());
      assertEquals(
          new dev.erst.gridgrind.core.model.layout.SheetDisplayReport(
              false, false, false, true, true),
          unfrozenLayout.presentation().display());
      assertEquals(
          Optional.of(CellColorReport.rgb(new RgbHex("#112233"))),
          unfrozenLayout.presentation().tabColor());
      assertEquals(
          new dev.erst.gridgrind.core.model.layout.SheetOutlineSummaryReport(false, false),
          unfrozenLayout.presentation().outlineSummary());
      assertEquals(
          new dev.erst.gridgrind.core.model.layout.SheetDefaultsReport(11, 18.5d),
          unfrozenLayout.presentation().sheetDefaults());
      assertEquals(
          List.of(
              new dev.erst.gridgrind.core.model.validation.IgnoredErrorReport(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B2"),
                  List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))),
          unfrozenLayout.presentation().ignoredErrors());
      assertEquals(3, unfrozenLayout.columns().size());
      assertEquals(3, unfrozenLayout.rows().size());

      sheet.layout().setPane(new PaneSpec.Frozen(1, 1, 1, 1));
      sheet.layout().setZoom(135);
      SheetLayoutReport frozenLayout = sheet.layout().snapshot();
      assertEquals(
          new dev.erst.gridgrind.core.model.layout.PaneReport.Frozen(1, 1, 1, 1),
          frozenLayout.pane());
      assertEquals(135, frozenLayout.zoomPercent());

      poiSheet.createSplitPane(2000, 2000, 0, 0, PaneType.LOWER_RIGHT);
      SheetLayoutReport splitLayout = sheet.layout().snapshot();
      assertEquals(
          new dev.erst.gridgrind.core.model.layout.PaneReport.Split(
              2000, 2000, 0, 0, ExcelPaneRegion.LOWER_RIGHT),
          splitLayout.pane());
    }
  }

  @Test
  void replacesHyperlinksOnRepeatedWrites() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet
          .annotations()
          .setHyperlink("F18", new HyperlinkTarget.Email("Report_Value@example.com"));
      sheet
          .annotations()
          .setHyperlink("F18", new HyperlinkTarget.Email("Summary.Total@example.com"));

      CellReport.BlankReport snapshot = (CellReport.BlankReport) sheet.cells().snapshotCell("F18");
      assertEquals(
          new HyperlinkTarget.Email("Summary.Total@example.com"),
          snapshot.hyperlink().orElseThrow());
    }
  }
}
