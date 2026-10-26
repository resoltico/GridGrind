package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.layout.PaneSpec;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelNamedRangeTargets;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import dev.erst.gridgrind.excel.ExcelTableSelection;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableSelection;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PaneType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;

/** Verifies that pre-save workbook expectations survive a reopened `.xlsx` package. */
final class XlsxRoundTripExpectationVerification {
  private XlsxRoundTripExpectationVerification() {}

  static void requireExpectedStyles(
      Map<String, Map<XlsxRoundTripExpectedStateSupport.CellCoordinate, CellStyleReport>>
          expectedStyles,
      XSSFWorkbook workbook) {
    if (expectedStyles.isEmpty()) {
      return;
    }
    for (Map.Entry<String, Map<XlsxRoundTripExpectedStateSupport.CellCoordinate, CellStyleReport>>
        sheetEntry : expectedStyles.entrySet()) {
      var sheet = workbook.getSheet(sheetEntry.getKey());
      if (sheet == null) {
        throw new IllegalStateException(
            "expected styled sheet must exist after round-trip: " + sheetEntry.getKey());
      }
      for (Map.Entry<XlsxRoundTripExpectedStateSupport.CellCoordinate, CellStyleReport> cellEntry :
          sheetEntry.getValue().entrySet()) {
        Row row = sheet.getRow(cellEntry.getKey().rowIndex());
        if (row == null) {
          throw new IllegalStateException(
              "expected styled row must exist after round-trip: "
                  + sheetEntry.getKey()
                  + "!"
                  + cellEntry.getKey().a1Address());
        }
        Cell cell = row.getCell(cellEntry.getKey().columnIndex());
        if (cell == null) {
          throw new IllegalStateException(
              "expected styled cell must exist after round-trip: "
                  + sheetEntry.getKey()
                  + "!"
                  + cellEntry.getKey().a1Address());
        }
        requireExpectedStyle(sheetEntry.getKey(), cellEntry.getKey(), cellEntry.getValue(), cell);
      }
    }
  }

  static void requireExpectedStyle(
      String sheetName,
      XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate,
      CellStyleReport expectedStyle,
      Cell cell) {
    CellStyleReport actualStyle =
        XlsxRoundTripVerifier.styleSnapshot(
            (XSSFWorkbook) cell.getSheet().getWorkbook(), (XSSFCellStyle) cell.getCellStyle());
    if (!expectedStyle.equals(actualStyle)) {
      throw new IllegalStateException(
          "style must survive .xlsx round-trip for %s!%s: expected %s but was %s"
              .formatted(sheetName, coordinate.a1Address(), expectedStyle, actualStyle));
    }
  }

  static void requireExpectedMetadata(
      Map<
              String,
              Map<
                  XlsxRoundTripExpectedStateSupport.CellCoordinate,
                  XlsxRoundTripExpectedStateSupport.ExpectedCellMetadata>>
          expectedMetadata,
      XSSFWorkbook workbook) {
    if (expectedMetadata.isEmpty()) {
      return;
    }
    for (Map.Entry<
            String,
            Map<
                XlsxRoundTripExpectedStateSupport.CellCoordinate,
                XlsxRoundTripExpectedStateSupport.ExpectedCellMetadata>>
        sheetEntry : expectedMetadata.entrySet()) {
      var sheet = workbook.getSheet(sheetEntry.getKey());
      if (sheet == null) {
        throw new IllegalStateException(
            "expected metadata sheet must exist after round-trip: " + sheetEntry.getKey());
      }
      for (Map.Entry<
              XlsxRoundTripExpectedStateSupport.CellCoordinate,
              XlsxRoundTripExpectedStateSupport.ExpectedCellMetadata>
          cellEntry : sheetEntry.getValue().entrySet()) {
        Row row = sheet.getRow(cellEntry.getKey().rowIndex());
        if (row == null) {
          throw new IllegalStateException(
              "expected metadata row must exist after round-trip: "
                  + sheetEntry.getKey()
                  + "!"
                  + cellEntry.getKey().a1Address());
        }
        Cell cell = row.getCell(cellEntry.getKey().columnIndex());
        if (cell == null) {
          throw new IllegalStateException(
              "expected metadata cell must exist after round-trip: "
                  + sheetEntry.getKey()
                  + "!"
                  + cellEntry.getKey().a1Address());
        }
        requireExpectedMetadata(
            sheetEntry.getKey(), cellEntry.getKey(), cellEntry.getValue(), cell);
      }
    }
  }

  static void requireExpectedMetadata(
      String sheetName,
      XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate,
      XlsxRoundTripExpectedStateSupport.ExpectedCellMetadata expectedMetadata,
      Cell cell) {
    requireEquals(
        sheetName,
        coordinate,
        "hyperlink",
        expectedMetadata.hyperlink(),
        XlsxRoundTripVerifier.hyperlink(cell));
    requireEquals(
        sheetName,
        coordinate,
        "comment",
        expectedMetadata.comment(),
        XlsxRoundTripVerifier.comment(cell));
  }

  static void requireExpectedNamedRanges(
      Map<
              XlsxRoundTripExpectedStateSupport.NamedRangeKey,
              XlsxRoundTripExpectedStateSupport.ExpectedNamedRange>
          expectedNamedRanges,
      XSSFWorkbook workbook) {
    if (expectedNamedRanges.isEmpty()) {
      return;
    }
    Map<
            XlsxRoundTripExpectedStateSupport.NamedRangeKey,
            XlsxRoundTripExpectedStateSupport.ExpectedNamedRange>
        actualNamedRanges = actualNamedRanges(workbook);
    for (Map.Entry<
            XlsxRoundTripExpectedStateSupport.NamedRangeKey,
            XlsxRoundTripExpectedStateSupport.ExpectedNamedRange>
        entry : expectedNamedRanges.entrySet()) {
      XlsxRoundTripExpectedStateSupport.ExpectedNamedRange actual =
          actualNamedRanges.get(entry.getKey());
      if (actual == null) {
        throw new IllegalStateException(
            "named range must survive .xlsx round-trip: " + entry.getKey().displayName());
      }
      if (!entry.getValue().equals(actual)) {
        throw new IllegalStateException(
            "named range must survive .xlsx round-trip for "
                + entry.getKey().displayName()
                + ": expected "
                + entry.getValue()
                + " but was "
                + actual);
      }
    }
  }

  static void requireExpectedDataValidations(
      Map<String, List<DataValidationEntryReport>> expectedDataValidations, Path workbookPath)
      throws IOException {
    if (expectedDataValidations.isEmpty()) {
      return;
    }
    try (ExcelWorkbook workbook = JazzerWorkbookIoSupport.openWorkbook(workbookPath)) {
      for (Map.Entry<String, List<DataValidationEntryReport>> entry :
          expectedDataValidations.entrySet()) {
        List<DataValidationEntryReport> actual =
            workbook
                .sheet(entry.getKey())
                .metadata()
                .dataValidations(new ExcelRangeSelection.All());
        if (!entry.getValue().equals(actual)) {
          throw new IllegalStateException(
              "data validations changed across round-trip for sheet "
                  + entry.getKey()
                  + ": expected "
                  + entry.getValue()
                  + " but was "
                  + actual);
        }
      }
    }
  }

  static void requireExpectedWorkbookState(
      XlsxRoundTripExpectedStateSupport.ExpectedWorkbookState expectedWorkbookState,
      Path workbookPath)
      throws IOException {
    try (ExcelWorkbook workbook = JazzerWorkbookIoSupport.openWorkbook(workbookPath)) {
      var actualWorkbookSummary =
          XlsxRoundTripExpectedStateSupport.expectedWorkbookSummary(workbook);
      if (!expectedWorkbookState.expectedWorkbookSummary().equals(actualWorkbookSummary)) {
        throw new IllegalStateException(
            "workbook summary changed across round-trip: expected "
                + expectedWorkbookState.expectedWorkbookSummary()
                + " but was "
                + actualWorkbookSummary);
      }
      for (Map.Entry<String, dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummary> entry :
          expectedWorkbookState.expectedSheetSummaries().entrySet()) {
        var actualSheetSummary =
            ((dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummaryResult)
                    new WorkbookExecutionEngine()
                        .read(
                            workbook,
                            new WorkbookReadCommand.GetSheetSummary(
                                new StepId("sheet-summary"), entry.getKey()))
                        .getFirst())
                .sheet();
        if (!entry.getValue().equals(actualSheetSummary)) {
          throw new IllegalStateException(
              "sheet summary changed across round-trip for "
                  + entry.getKey()
                  + ": expected "
                  + entry.getValue()
                  + " but was "
                  + actualSheetSummary);
        }
      }
      for (Map.Entry<
              String,
              Map<XlsxRoundTripExpectedStateSupport.CellCoordinate, List<RichTextRunReport>>>
          entry : expectedWorkbookState.expectedRichText().entrySet()) {
        for (Map.Entry<XlsxRoundTripExpectedStateSupport.CellCoordinate, List<RichTextRunReport>>
            cellEntry : entry.getValue().entrySet()) {
          CellReport actualSnapshot =
              workbook.sheet(entry.getKey()).cells().snapshotCell(cellEntry.getKey().a1Address());
          if (!(actualSnapshot instanceof CellReport.TextReport textSnapshot)) {
            throw new IllegalStateException(
                "rich text cell must reopen as a text snapshot for "
                    + entry.getKey()
                    + "!"
                    + cellEntry.getKey().a1Address());
          }
          requireEquals(
              entry.getKey(),
              cellEntry.getKey(),
              "textValue",
              cellEntry.getValue().stream()
                  .map(RichTextRunReport::text)
                  .collect(java.util.stream.Collectors.joining()),
              textSnapshot.value().orElseThrow().value());
          requireEquals(
              entry.getKey(),
              cellEntry.getKey(),
              "richText",
              cellEntry.getValue(),
              textSnapshot.runs().orElseThrow());
        }
      }
      for (Map.Entry<String, List<DrawingObjectReport>> entry :
          expectedWorkbookState.expectedDrawingObjects().entrySet()) {
        List<DrawingObjectReport> actualDrawingObjects =
            workbook.sheet(entry.getKey()).drawings().drawingObjects();
        if (!entry.getValue().equals(actualDrawingObjects)) {
          throw new IllegalStateException(
              "drawing objects changed across round-trip for "
                  + entry.getKey()
                  + ": expected "
                  + entry.getValue()
                  + " but was "
                  + actualDrawingObjects);
        }
      }
      for (Map.Entry<String, List<ChartReport>> entry :
          expectedWorkbookState.expectedCharts().entrySet()) {
        List<ChartReport> actualCharts = workbook.sheet(entry.getKey()).drawings().charts();
        if (!entry.getValue().equals(actualCharts)) {
          throw new IllegalStateException(
              "charts changed across round-trip for "
                  + entry.getKey()
                  + ": expected "
                  + entry.getValue()
                  + " but was "
                  + actualCharts);
        }
      }
      List<PivotTableReport> actualPivots =
          ((dev.erst.gridgrind.excel.WorkbookDrawingResult.PivotTablesResult)
                  new WorkbookExecutionEngine()
                      .read(
                          workbook,
                          new WorkbookReadCommand.GetPivotTables(
                              new StepId("pivots"), new ExcelPivotTableSelection.All()))
                      .getFirst())
              .pivotTables();
      if (!expectedWorkbookState.expectedPivots().equals(actualPivots)) {
        throw new IllegalStateException(
            "pivot tables changed across round-trip: expected "
                + expectedWorkbookState.expectedPivots()
                + " but was "
                + actualPivots);
      }
    }
  }

  static void requireExpectedSheetLayouts(
      Map<String, XlsxRoundTripExpectedStateSupport.ExpectedSheetLayoutState> expectedSheetLayouts,
      Path workbookPath)
      throws IOException {
    if (expectedSheetLayouts.isEmpty()) {
      return;
    }
    try (InputStream inputStream = Files.newInputStream(workbookPath);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {
      for (Map.Entry<String, XlsxRoundTripExpectedStateSupport.ExpectedSheetLayoutState> entry :
          expectedSheetLayouts.entrySet()) {
        var sheet = workbook.getSheet(entry.getKey());
        if (sheet == null) {
          throw new IllegalStateException(
              "expected sheet layout sheet must exist after round-trip: " + entry.getKey());
        }
        requireExpectedSheetLayout(entry.getKey(), entry.getValue(), sheet);
      }
    }
  }

  static void requireExpectedSheetLayout(
      String sheetName,
      XlsxRoundTripExpectedStateSupport.ExpectedSheetLayoutState expectedSheetLayout,
      XSSFSheet sheet) {
    PaneSpec actualPane = pane(sheet);
    int actualZoomPercent = zoomPercent(sheet);
    if (expectedSheetLayout.pane() != null && !expectedSheetLayout.pane().equals(actualPane)) {
      throw new IllegalStateException(
          "sheet pane changed across round-trip for "
              + sheetName
              + ": expected "
              + expectedSheetLayout.pane()
              + " but was "
              + actualPane);
    }
    if (expectedSheetLayout.zoomPercent() != null
        && !expectedSheetLayout.zoomPercent().equals(actualZoomPercent)) {
      throw new IllegalStateException(
          "sheet zoom changed across round-trip for "
              + sheetName
              + ": expected "
              + expectedSheetLayout.zoomPercent()
              + " but was "
              + actualZoomPercent);
    }
    for (Map.Entry<Integer, XlsxRoundTripExpectedStateSupport.ExpectedRowLayoutState> rowEntry :
        expectedSheetLayout.rows().entrySet()) {
      requireExpectedRowLayout(sheetName, rowEntry.getKey(), rowEntry.getValue(), sheet);
    }
    for (Map.Entry<Integer, XlsxRoundTripExpectedStateSupport.ExpectedColumnLayoutState>
        columnEntry : expectedSheetLayout.columns().entrySet()) {
      requireExpectedColumnLayout(sheetName, columnEntry.getKey(), columnEntry.getValue(), sheet);
    }
  }

  static void requireExpectedRowLayout(
      String sheetName,
      int rowIndex,
      XlsxRoundTripExpectedStateSupport.ExpectedRowLayoutState expectedRowLayout,
      XSSFSheet sheet) {
    Row row = sheet.getRow(rowIndex);
    Boolean hidden =
        row instanceof org.apache.poi.xssf.usermodel.XSSFRow xssfRow && xssfRow.getZeroHeight();
    Integer outlineLevel = row == null ? 0 : Math.max(0, row.getOutlineLevel());
    Boolean collapsed =
        row instanceof org.apache.poi.xssf.usermodel.XSSFRow xssfRow
            && xssfRow.getCTRow().getCollapsed();
    requireLayoutField(sheetName, "row", rowIndex, "hidden", expectedRowLayout.hidden(), hidden);
    requireLayoutField(
        sheetName, "row", rowIndex, "outlineLevel", expectedRowLayout.outlineLevel(), outlineLevel);
    requireLayoutField(
        sheetName, "row", rowIndex, "collapsed", expectedRowLayout.collapsed(), collapsed);
  }

  static void requireExpectedColumnLayout(
      String sheetName,
      int columnIndex,
      XlsxRoundTripExpectedStateSupport.ExpectedColumnLayoutState expectedColumnLayout,
      XSSFSheet sheet) {
    Boolean hidden = sheet.isColumnHidden(columnIndex);
    Integer outlineLevel = sheet.getColumnOutlineLevel(columnIndex);
    Boolean collapsed = columnCollapsed(sheet, columnIndex);
    requireLayoutField(
        sheetName, "column", columnIndex, "hidden", expectedColumnLayout.hidden(), hidden);
    requireLayoutField(
        sheetName,
        "column",
        columnIndex,
        "outlineLevel",
        expectedColumnLayout.outlineLevel(),
        outlineLevel);
    requireLayoutField(
        sheetName, "column", columnIndex, "collapsed", expectedColumnLayout.collapsed(), collapsed);
  }

  static void requireExpectedConditionalFormatting(
      Map<String, List<ConditionalFormattingEntryReport>> expectedConditionalFormatting,
      Path workbookPath)
      throws IOException {
    if (expectedConditionalFormatting.isEmpty()) {
      return;
    }
    try (ExcelWorkbook workbook = JazzerWorkbookIoSupport.openWorkbook(workbookPath)) {
      WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
      for (Map.Entry<String, List<ConditionalFormattingEntryReport>> entry :
          expectedConditionalFormatting.entrySet()) {
        var actual =
            ((dev.erst.gridgrind.excel.WorkbookRuleResult.ConditionalFormattingResult)
                    readExecutor
                        .read(
                            workbook,
                            new WorkbookReadCommand.GetConditionalFormatting(
                                new StepId("conditionalFormatting"),
                                entry.getKey(),
                                new ExcelRangeSelection.All()))
                        .getFirst())
                .conditionalFormattingBlocks();
        if (!entry.getValue().equals(actual)) {
          throw new IllegalStateException(
              "conditional formatting changed across round-trip for sheet "
                  + entry.getKey()
                  + ": expected "
                  + entry.getValue()
                  + " but was "
                  + actual);
        }
      }
    }
  }

  static void requireExpectedAutofilters(
      Map<String, List<AutofilterEntryReport>> expectedAutofilters, Path workbookPath)
      throws IOException {
    if (expectedAutofilters.isEmpty()) {
      return;
    }
    try (ExcelWorkbook workbook = JazzerWorkbookIoSupport.openWorkbook(workbookPath)) {
      WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
      for (Map.Entry<String, List<AutofilterEntryReport>> entry : expectedAutofilters.entrySet()) {
        var actual =
            ((dev.erst.gridgrind.excel.WorkbookRuleResult.AutofiltersResult)
                    readExecutor
                        .read(
                            workbook,
                            new WorkbookReadCommand.GetAutofilters(
                                new StepId("autofilters"), entry.getKey()))
                        .getFirst())
                .autofilters();
        if (!entry.getValue().equals(actual)) {
          throw new IllegalStateException(
              "autofilters changed across round-trip for sheet "
                  + entry.getKey()
                  + ": expected "
                  + entry.getValue()
                  + " but was "
                  + actual);
        }
      }
    }
  }

  static void requireExpectedTables(List<TableEntryReport> expectedTables, Path workbookPath)
      throws IOException {
    if (expectedTables.isEmpty()) {
      return;
    }
    try (ExcelWorkbook workbook = JazzerWorkbookIoSupport.openWorkbook(workbookPath)) {
      WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
      var actual =
          ((dev.erst.gridgrind.excel.WorkbookRuleResult.TablesResult)
                  readExecutor
                      .read(
                          workbook,
                          new WorkbookReadCommand.GetTables(
                              new StepId("tables"), new ExcelTableSelection.All()))
                      .getFirst())
              .tables();
      if (!expectedTables.equals(actual)) {
        throw new IllegalStateException(
            "tables changed across round-trip: expected " + expectedTables + " but was " + actual);
      }
    }
  }

  static void requireLayoutField(
      String sheetName, String axis, int index, String fieldName, Object expected, Object actual) {
    if (expected == null || expected.equals(actual)) {
      return;
    }
    throw new IllegalStateException(
        "sheet layout %s %d field %s changed across round-trip for %s: expected %s but was %s"
            .formatted(axis, index, fieldName, sheetName, expected, actual));
  }

  static PaneSpec pane(XSSFSheet sheet) {
    PaneInformation paneInformation = sheet.getPaneInformation();
    if (paneInformation == null) {
      return new PaneSpec.None();
    }
    if (paneInformation.isFreezePane()) {
      return new PaneSpec.Frozen(
          paneInformation.getVerticalSplitPosition(),
          paneInformation.getHorizontalSplitPosition(),
          paneInformation.getVerticalSplitLeftColumn(),
          paneInformation.getHorizontalSplitTopRow());
    }
    return new PaneSpec.Split(
        paneInformation.getVerticalSplitPosition(),
        paneInformation.getHorizontalSplitPosition(),
        paneInformation.getVerticalSplitLeftColumn(),
        paneInformation.getHorizontalSplitTopRow(),
        paneRegion(paneInformation.getActivePaneType()));
  }

  static int zoomPercent(XSSFSheet sheet) {
    var sheetView = sheet.getCTWorksheet().getSheetViews().getSheetViewArray(0);
    return sheetView.isSetZoomScale() ? Math.toIntExact(sheetView.getZoomScale()) : 100;
  }

  static dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion paneRegion(PaneType paneType) {
    return switch (paneType) {
      case UPPER_LEFT -> dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion.UPPER_LEFT;
      case UPPER_RIGHT -> dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion.UPPER_RIGHT;
      case LOWER_LEFT -> dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion.LOWER_LEFT;
      case LOWER_RIGHT -> dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion.LOWER_RIGHT;
    };
  }

  static boolean columnCollapsed(XSSFSheet sheet, int columnIndex) {
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        if (columnIndex + 1 >= col.getMin() && columnIndex + 1 <= col.getMax()) {
          return col.getCollapsed();
        }
      }
    }
    return false;
  }

  static Map<
          XlsxRoundTripExpectedStateSupport.NamedRangeKey,
          XlsxRoundTripExpectedStateSupport.ExpectedNamedRange>
      actualNamedRanges(XSSFWorkbook workbook) {
    LinkedHashMap<
            XlsxRoundTripExpectedStateSupport.NamedRangeKey,
            XlsxRoundTripExpectedStateSupport.ExpectedNamedRange>
        actualNamedRanges = new LinkedHashMap<>();
    for (Name name : workbook.getAllNames()) {
      if (!shouldExpose(name)) {
        continue;
      }
      NamedRangeScopeReport scope = toScope(workbook, name.getSheetIndex());
      String refersToFormula = Objects.requireNonNullElse(name.getRefersToFormula(), "");
      actualNamedRanges.put(
          new XlsxRoundTripExpectedStateSupport.NamedRangeKey(name.getNameName(), scope),
          new XlsxRoundTripExpectedStateSupport.ExpectedNamedRange(
              scope,
              refersToFormula,
              ExcelNamedRangeTargets.resolveTarget(refersToFormula, scope).orElse(null)));
    }
    return Map.copyOf(actualNamedRanges);
  }

  static boolean shouldExpose(Name name) {
    String nameName = name.getNameName();
    return !name.isFunctionName()
        && !name.isHidden()
        && nameName != null
        && !nameName.startsWith("_xlnm.")
        && !nameName.startsWith("_XLNM.");
  }

  static NamedRangeScopeReport toScope(XSSFWorkbook workbook, int sheetIndex) {
    if (sheetIndex < 0) {
      return new NamedRangeScopeReport.Workbook();
    }
    return new NamedRangeScopeReport.Sheet(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(workbook.getSheetName(sheetIndex)));
  }

  static void requireEquals(
      String sheetName,
      XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate,
      String fieldName,
      @Nullable Object expected,
      @Nullable Object actual) {
    if (expected == null) {
      return;
    }
    if (!expected.equals(actual)) {
      throw new IllegalStateException(
          "%s must survive .xlsx round-trip for %s!%s: expected %s but was %s"
              .formatted(fieldName, sheetName, coordinate.a1Address(), expected, actual));
    }
  }
}
