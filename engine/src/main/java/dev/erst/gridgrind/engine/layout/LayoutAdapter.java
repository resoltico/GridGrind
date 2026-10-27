package dev.erst.gridgrind.engine.layout;

import dev.erst.gridgrind.core.model.layout.ColumnLayoutReport;
import dev.erst.gridgrind.core.model.layout.MergedRegionReport;
import dev.erst.gridgrind.core.model.layout.PaneSpec;
import dev.erst.gridgrind.core.model.layout.RowLayoutReport;
import dev.erst.gridgrind.core.model.layout.SheetDefaultsReport;
import dev.erst.gridgrind.core.model.layout.SheetDefaultsSpec;
import dev.erst.gridgrind.core.model.layout.SheetDisplayReport;
import dev.erst.gridgrind.core.model.layout.SheetDisplaySpec;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummaryReport;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummarySpec;
import dev.erst.gridgrind.core.model.layout.SheetPresentationReport;
import dev.erst.gridgrind.core.model.layout.SheetPresentationSpec;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorReport;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion;
import dev.erst.gridgrind.core.vocabulary.ExcelRowSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits;
import dev.erst.gridgrind.engine.comments.CommentAdapter;
import dev.erst.gridgrind.engine.poi.CellBounds;
import dev.erst.gridgrind.engine.poi.SheetStructure;
import dev.erst.gridgrind.engine.print.PrintAdapter;
import dev.erst.gridgrind.engine.styles.ColorAdapter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.EnumSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IgnoredErrorType;
import org.apache.poi.ss.usermodel.PaneType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetView;

/** Owns sheet layout, panes, display/defaults, merged regions and row/column outline facts. */
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired"})
public final class LayoutAdapter {
  private final Sheet sheet;
  private final PrintAdapter printAdapter;

  public LayoutAdapter(Sheet sheet, PrintAdapter printAdapter) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.printAdapter = Objects.requireNonNull(printAdapter, "printAdapter must not be null");
  }

  /** Returns factual presentation, pane, dimensions and outline state for the bound sheet. */
  public SheetLayoutReport snapshot() {
    return new SheetLayoutReport(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheet.getSheetName()),
        pane(xssfSheet()),
        zoomPercent(xssfSheet()),
        presentation(xssfSheet()),
        columnLayouts(xssfSheet()),
        rowLayouts(xssfSheet()));
  }

  public LayoutAdapter mergeCells(String range) {
    CellBounds excelRange = CellBounds.parse(range);
    requireMergeableRange(range, excelRange);
    if (SheetStructure.findMergedRegionIndex(sheet, excelRange) >= 0) {
      return this;
    }
    SheetStructure.requireNoMergedRegionOverlap(sheet, excelRange);
    sheet.addMergedRegion(SheetStructure.toCellRangeAddress(excelRange));
    return this;
  }

  public LayoutAdapter unmergeCells(String range) {
    CellBounds excelRange = CellBounds.parse(range);
    int mergedRegionIndex = SheetStructure.findMergedRegionIndex(sheet, excelRange);
    if (mergedRegionIndex < 0) {
      throw new IllegalArgumentException("No merged region matches range: " + range);
    }
    sheet.removeMergedRegion(mergedRegionIndex);
    return this;
  }

  public LayoutAdapter setPane(PaneSpec pane) {
    Objects.requireNonNull(pane, "pane must not be null");
    setPane(xssfSheet(), pane);
    return this;
  }

  public LayoutAdapter setZoom(int zoomPercent) {
    requireZoomPercent(zoomPercent);
    setZoomPercent(xssfSheet(), zoomPercent);
    return this;
  }

  public LayoutAdapter setPresentation(SheetPresentationSpec presentation) {
    Objects.requireNonNull(presentation, "presentation must not be null");
    setPresentation(xssfSheet(), presentation);
    return this;
  }

  public LayoutAdapter setPrintLayout(ResolvedPrintLayout printLayout) {
    Objects.requireNonNull(printLayout, "printLayout must not be null");
    printAdapter.setPrintLayout(xssfSheet(), printLayout);
    return this;
  }

  public LayoutAdapter clearPrintLayout() {
    printAdapter.clearPrintLayout(xssfSheet());
    return this;
  }

  public List<MergedRegionReport> mergedRegions() {
    List<MergedRegionReport> mergedRegions = new ArrayList<>(sheet.getNumMergedRegions());
    for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
      mergedRegions.add(
          new MergedRegionReport(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel(
                  sheet.getMergedRegion(regionIndex).formatAsString())));
    }
    return List.copyOf(mergedRegions);
  }

  public ResolvedPrintLayout printLayout() {
    return printAdapter.printLayout(xssfSheet());
  }

  public PrintLayoutReport printLayoutSnapshot() {
    return printAdapter.printLayoutSnapshot(xssfSheet());
  }

  private XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
  }

  private static void requireMergeableRange(String range, CellBounds excelRange) {
    if (excelRange.rowCount() == 1 && excelRange.columnCount() == 1) {
      throw new IllegalArgumentException("range must span at least two cells: " + range);
    }
  }

  /** Applies the provided presentation payload authoritatively to the target sheet. */
  public static void setPresentation(XSSFSheet sheet, SheetPresentationSpec presentation) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(presentation, "presentation must not be null");

    applyDisplay(sheet, presentation.display());
    applyTabColor(sheet, presentation.tabColor());
    applyOutlineSummary(sheet, presentation.outlineSummary());
    applyDefaults(sheet, presentation.sheetDefaults());
    applyIgnoredErrors(sheet, presentation.ignoredErrors());
  }

  /** Clears sheet presentation back to the effective workbook defaults. */
  public static void clearPresentation(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    setPresentation(sheet, SheetPresentationSpec.defaults());
  }

  /** Returns the factual sheet-presentation state currently stored on the provided sheet. */
  public static SheetPresentationReport presentation(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return new SheetPresentationReport(
        display(sheet),
        tabColor(sheet),
        outlineSummary(sheet),
        defaults(sheet),
        ignoredErrors(sheet));
  }

  private static void applyDisplay(XSSFSheet sheet, SheetDisplaySpec display) {
    sheet.setDisplayGridlines(display.displayGridlines());
    sheet.setDisplayZeros(display.displayZeros());
    sheet.setDisplayRowColHeadings(display.displayRowColHeadings());
    sheet.setDisplayFormulas(display.displayFormulas());
    sheet.setRightToLeft(display.rightToLeft());
  }

  private static SheetDisplayReport display(XSSFSheet sheet) {
    return new SheetDisplayReport(
        sheet.isDisplayGridlines(),
        sheet.isDisplayZeros(),
        sheet.isDisplayRowColHeadings(),
        sheet.isDisplayFormulas(),
        sheet.isRightToLeft());
  }

  private static void applyTabColor(XSSFSheet sheet, Optional<ColorSpec> tabColor) {
    Objects.requireNonNull(tabColor, "tabColor must not be null");
    if (tabColor.isEmpty()) {
      clearTabColor(sheet);
      return;
    }
    sheet.setTabColor(ColorAdapter.toXssfColor(sheet.getWorkbook(), tabColor.orElseThrow()));
  }

  private static Optional<CellColorReport> tabColor(XSSFSheet sheet) {
    XSSFColor color = sheet.getTabColor();
    return ColorAdapter.snapshot(color);
  }

  public static void clearTabColor(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetSheetPr()) {
      return;
    }
    var sheetPr = sheet.getCTWorksheet().getSheetPr();
    if (sheetPr.isSetTabColor()) {
      sheetPr.unsetTabColor();
    }
  }

  private static void applyOutlineSummary(XSSFSheet sheet, SheetOutlineSummarySpec outlineSummary) {
    sheet.setRowSumsBelow(outlineSummary.rowSumsBelow());
    sheet.setRowSumsRight(outlineSummary.rowSumsRight());
  }

  private static SheetOutlineSummaryReport outlineSummary(XSSFSheet sheet) {
    return new SheetOutlineSummaryReport(sheet.getRowSumsBelow(), sheet.getRowSumsRight());
  }

  private static void applyDefaults(XSSFSheet sheet, SheetDefaultsSpec defaults) {
    ExcelSheetLayoutLimits.requireDefaultColumnWidth(
        defaults.defaultColumnWidth(), "defaultColumnWidth");
    ExcelSheetLayoutLimits.requireRowHeightPoints(
        defaults.defaultRowHeightPoints(), "defaultRowHeightPoints");
    sheet.setDefaultColumnWidth(defaults.defaultColumnWidth());
    sheet.setDefaultRowHeightInPoints((float) defaults.defaultRowHeightPoints());
  }

  private static SheetDefaultsReport defaults(XSSFSheet sheet) {
    double defaultRowHeightPoints = sheet.getDefaultRowHeightInPoints();
    return new SheetDefaultsReport(
        sheet.getDefaultColumnWidth(),
        defaultRowHeightPoints <= 0.0d
            ? SheetDefaultsReport.defaults().defaultRowHeightPoints()
            : defaultRowHeightPoints);
  }

  private static void applyIgnoredErrors(XSSFSheet sheet, List<IgnoredErrorSpec> ignoredErrors) {
    clearIgnoredErrors(sheet);
    for (IgnoredErrorSpec ignoredError : ignoredErrors) {
      sheet.addIgnoredErrors(
          SheetStructure.toCellRangeAddress(CellBounds.parse(ignoredError.range().value())),
          ignoredError.errorTypes().stream()
              .map(LayoutAdapter::toPoiIgnoredError)
              .toArray(IgnoredErrorType[]::new));
    }
  }

  private static void clearIgnoredErrors(XSSFSheet sheet) {
    if (sheet.getCTWorksheet().isSetIgnoredErrors()) {
      sheet.getCTWorksheet().unsetIgnoredErrors();
    }
  }

  @SuppressWarnings({"PMD.LooseCoupling", "PMD.UseConcurrentHashMap"})
  private static List<IgnoredErrorReport> ignoredErrors(XSSFSheet sheet) {
    Map<String, EnumSet<ExcelIgnoredErrorType>> ignoredErrorsByRange = new TreeMap<>();
    for (Map.Entry<IgnoredErrorType, java.util.Set<CellRangeAddress>> entry :
        sheet.getIgnoredErrors().entrySet()) {
      ExcelIgnoredErrorType errorType = fromPoiIgnoredError(entry.getKey());
      entry.getValue().stream()
          .sorted(Comparator.comparing(CellRangeAddress::formatAsString))
          .map(CellRangeAddress::formatAsString)
          .forEach(
              range ->
                  ignoredErrorsByRange
                      .computeIfAbsent(
                          range, ignoredError -> EnumSet.noneOf(ExcelIgnoredErrorType.class))
                      .add(errorType));
    }
    return ignoredErrorsByRange.entrySet().stream()
        .map(
            entry ->
                new IgnoredErrorReport(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel(entry.getKey()),
                    List.copyOf(entry.getValue())))
        .toList();
  }

  /** Applies one explicit pane state to the provided sheet. */
  public static void setPane(XSSFSheet sheet, PaneSpec pane) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(pane, "pane must not be null");
    switch (pane) {
      case PaneSpec.None _ -> sheet.createFreezePane(0, 0, 0, 0);
      case PaneSpec.Frozen frozen ->
          sheet.createFreezePane(
              frozen.splitColumn(), frozen.splitRow(), frozen.leftmostColumn(), frozen.topRow());
      case PaneSpec.Split split ->
          sheet.createSplitPane(
              split.xSplitPosition(),
              split.ySplitPosition(),
              split.leftmostColumn(),
              split.topRow(),
              toPoiPane(split.activePane()));
    }
  }

  /** Returns the current pane state captured on the provided sheet. */
  public static dev.erst.gridgrind.core.model.layout.PaneReport pane(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    PaneInformation paneInformation = sheet.getPaneInformation();
    if (paneInformation == null) {
      return new dev.erst.gridgrind.core.model.layout.PaneReport.None();
    }
    if (paneInformation.isFreezePane()) {
      return new dev.erst.gridgrind.core.model.layout.PaneReport.Frozen(
          paneInformation.getVerticalSplitPosition(),
          paneInformation.getHorizontalSplitPosition(),
          paneInformation.getVerticalSplitLeftColumn(),
          paneInformation.getHorizontalSplitTopRow());
    }
    return new dev.erst.gridgrind.core.model.layout.PaneReport.Split(
        paneInformation.getVerticalSplitPosition(),
        paneInformation.getHorizontalSplitPosition(),
        paneInformation.getVerticalSplitLeftColumn(),
        paneInformation.getHorizontalSplitTopRow(),
        fromPoiPane(paneInformation.getActivePaneType()));
  }

  /** Applies one zoom percentage to the provided sheet. */
  private static void setZoomPercent(XSSFSheet sheet, int zoomPercent) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireZoomPercent(zoomPercent);
    sheet.setZoom(zoomPercent);
  }

  /** Returns the effective zoom percentage for the provided sheet. */
  public static int zoomPercent(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    CTSheetView sheetView = sheet.getCTWorksheet().getSheetViews().getSheetViewArray(0);
    return sheetView.isSetZoomScale() ? Math.toIntExact(sheetView.getZoomScale()) : 100;
  }

  /** Validates one requested zoom percentage. */
  public static void requireZoomPercent(int zoomPercent) { // LIM-022
    ExcelSheetLayoutLimits.requireZoomPercent(zoomPercent, "zoomPercent");
  }

  public static ExcelPaneRegion fromPoiPane(PaneType paneType) {
    Objects.requireNonNull(paneType, "paneType must not be null");
    return ExcelPaneRegion.valueOf(paneType.name());
  }

  public static PaneType toPoiPane(ExcelPaneRegion paneRegion) {
    Objects.requireNonNull(paneRegion, "paneRegion must not be null");
    return PaneType.valueOf(paneRegion.name());
  }

  public static IgnoredErrorType toPoiIgnoredError(ExcelIgnoredErrorType ignoredErrorType) {
    Objects.requireNonNull(ignoredErrorType, "ignoredErrorType must not be null");
    return IgnoredErrorType.valueOf(ignoredErrorType.name());
  }

  public static ExcelIgnoredErrorType fromPoiIgnoredError(IgnoredErrorType ignoredErrorType) {
    Objects.requireNonNull(ignoredErrorType, "ignoredErrorType must not be null");
    return ExcelIgnoredErrorType.valueOf(ignoredErrorType.name());
  }

  public static List<RowLayoutReport> rowLayouts(XSSFSheet sheet) {
    int lastRowIndex = sheet.getLastRowNum();
    if (lastRowIndex < 0) {
      return List.of();
    }
    List<RowLayoutReport> rows = new ArrayList<>(lastRowIndex + 1);
    for (int rowIndex = 0; rowIndex <= lastRowIndex; rowIndex++) {
      Row row = sheet.getRow(rowIndex);
      double heightPoints =
          row == null ? sheet.getDefaultRowHeight() / 20.0d : row.getHeight() / 20.0d;
      rows.add(
          new RowLayoutReport(
              rowIndex,
              heightPoints,
              row instanceof XSSFRow xssfRow && xssfRow.getZeroHeight(),
              row == null ? 0 : Math.max(0, row.getOutlineLevel()),
              row instanceof XSSFRow xssfRow && xssfRow.getCTRow().getCollapsed()));
    }
    return List.copyOf(rows);
  }

  public static void ensureRowsExist(XSSFSheet sheet, ExcelRowSpan rows) {
    for (int rowIndex = rows.firstRowIndex(); rowIndex <= rows.lastRowIndex(); rowIndex++) {
      if (sheet.getRow(rowIndex) == null) {
        sheet.createRow(rowIndex);
      }
    }
  }

  public static void expandRows(XSSFSheet sheet, ExcelRowSpan rows) {
    for (int rowIndex = rows.firstRowIndex(); rowIndex <= rows.lastRowIndex(); rowIndex++) {
      Row row = sheet.getRow(rowIndex);
      row.setZeroHeight(false);
    }
    clearExpandedGroupControlRow(sheet, rows);
  }

  public static boolean isRowGroupCollapsed(XSSFSheet sheet, ExcelRowSpan rows) {
    if (rows.lastRowIndex() >= ExcelRowSpan.MAX_ROW_INDEX) {
      return false;
    }
    Row controlRow = sheet.getRow(rows.lastRowIndex() + 1);
    return controlRow instanceof XSSFRow xssfRow && xssfRow.getCTRow().getCollapsed();
  }

  public static void clearExpandedGroupControlRow(XSSFSheet sheet, ExcelRowSpan rows) {
    if (rows.lastRowIndex() < ExcelRowSpan.MAX_ROW_INDEX) {
      Row controlRow = sheet.getRow(rows.lastRowIndex() + 1);
      if (controlRow instanceof XSSFRow xssfRow) {
        xssfRow.getCTRow().setCollapsed(false);
      }
    }
  }

  public static void collapseRows(XSSFSheet sheet, ExcelRowSpan rows) {
    XSSFRow controlRow = null;
    if (rows.lastRowIndex() < ExcelRowSpan.MAX_ROW_INDEX) {
      Row existingControlRow = sheet.getRow(rows.lastRowIndex() + 1);
      controlRow =
          existingControlRow instanceof XSSFRow xssfRow
              ? xssfRow
              : (XSSFRow) sheet.createRow(rows.lastRowIndex() + 1);
      controlRow.setZeroHeight(false);
    }
    sheet.setRowGroupCollapsed(rows.firstRowIndex(), true);
    if (controlRow != null) {
      controlRow.getCTRow().setCollapsed(true);
    }
  }

  public static void prepareColumnsForUngroup(XSSFSheet sheet, ExcelColumnSpan columns) {
    for (int columnIndex = columns.firstColumnIndex();
        columnIndex <= columns.lastColumnIndex();
        columnIndex++) {
      sheet.setColumnHidden(columnIndex, false);
    }
    clearExpandedGroupControlColumn(sheet, columns);
  }

  public static void collapseColumns(XSSFSheet sheet, ExcelColumnSpan columns) {
    sheet.setColumnGroupCollapsed(columns.firstColumnIndex(), true);
  }

  public static void clearExpandedGroupControlColumn(XSSFSheet sheet, ExcelColumnSpan columns) {
    if (columns.lastColumnIndex() < ExcelColumnSpan.MAX_COLUMN_INDEX) {
      setColumnCollapsed(sheet, columns.lastColumnIndex() + 1, false);
    }
  }

  public static void clearTrailingCells(
      XSSFSheet sheet, int firstColumnIndex, int lastColumnIndex) {
    for (Row row : sheet) {
      for (int columnIndex = firstColumnIndex; columnIndex <= lastColumnIndex; columnIndex++) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
          continue;
        }
        cell.removeHyperlink();
        CommentAdapter.clearCellComment(cell);
        row.removeCell(cell);
      }
    }
  }

  public static Map<Integer, CTCol> mutableColumnDefinitionsCopy(
      Map<Integer, CTCol> explicitColumns) {
    return new LinkedHashMap<>(explicitColumns);
  }

  /** Returns the last column index implied by cells or explicit column metadata. */
  public static int lastColumnIndex(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    int lastColumnIndex = -1;
    for (Row row : sheet) {
      lastColumnIndex = Math.max(lastColumnIndex, row.getLastCellNum() - 1);
    }
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        lastColumnIndex = Math.max(lastColumnIndex, (int) col.getMax() - 1);
      }
    }
    return lastColumnIndex;
  }

  /** Returns the sheet column layouts including hidden, outline, and collapsed state. */
  public static List<ColumnLayoutReport> columnLayouts(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    int lastColumnIndex = lastColumnIndex(sheet);
    if (lastColumnIndex < 0) {
      return List.of();
    }
    Map<Integer, CTCol> effectiveColumns = effectiveColumnDefinitions(sheet);
    List<ColumnLayoutReport> columns = new ArrayList<>(lastColumnIndex + 1);
    for (int columnIndex = 0; columnIndex <= lastColumnIndex; columnIndex++) {
      CTCol columnDefinition = effectiveColumns.get(columnIndex);
      columns.add(
          new ColumnLayoutReport(
              columnIndex,
              sheet.getColumnWidth(columnIndex) / 256.0d,
              columnDefinition != null && columnDefinition.getHidden(),
              columnDefinition == null ? 0 : (int) columnDefinition.getOutlineLevel(),
              columnDefinition != null && columnDefinition.getCollapsed()));
    }
    return List.copyOf(columns);
  }

  /** Applies a collapsed marker to the explicit column definition that owns the target index. */
  public static void setColumnCollapsed(XSSFSheet sheet, int columnIndex, boolean collapsed) {
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        if (columnIndex + 1 >= col.getMin() && columnIndex + 1 <= col.getMax()) {
          col.setCollapsed(collapsed);
        }
      }
    }
  }

  public static void normalizeColumnDefinitionContainer(XSSFSheet sheet) {
    if (!requiresColumnDefinitionCanonicalization(sheet)) {
      return;
    }
    canonicalizeColumnDefinitions(sheet);
  }

  public static void canonicalizeColumnDefinitions(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    rebuildColumnDefinitions(sheet, effectiveColumnDefinitions(sheet));
  }

  public static void rebuildColumnDefinitions(
      XSSFSheet sheet, Map<Integer, CTCol> explicitColumns) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(explicitColumns, "explicitColumns must not be null");
    rebuildColumnDefinitionsInternal(sheet, explicitColumns);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  public static Map<Integer, CTCol> snapshotColumnDefinitions(XSSFSheet sheet) {
    Map<Integer, CTCol> explicitColumns = new LinkedHashMap<>();
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        for (int columnIndex = (int) col.getMin() - 1;
            columnIndex <= (int) col.getMax() - 1;
            columnIndex++) {
          CTCol definition = copyOf(col);
          definition.setMin(columnIndex + 1L);
          definition.setMax(columnIndex + 1L);
          explicitColumns.put(columnIndex, definition);
        }
      }
    }
    return Map.copyOf(explicitColumns);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  public static Map<Integer, CTCol> shiftedForInsert(
      Map<Integer, CTCol> explicitColumns, int columnIndex, int columnCount) {
    Map<Integer, CTCol> shifted = new LinkedHashMap<>();
    explicitColumns.forEach(
        (existingColumnIndex, definition) ->
            shifted.put(
                existingColumnIndex >= columnIndex
                    ? existingColumnIndex + columnCount
                    : existingColumnIndex,
                definition));
    return Map.copyOf(shifted);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  public static Map<Integer, CTCol> shiftedForDelete(
      Map<Integer, CTCol> explicitColumns, ExcelColumnSpan columns) {
    Map<Integer, CTCol> shifted = new LinkedHashMap<>();
    explicitColumns.forEach(
        (existingColumnIndex, definition) -> {
          if (existingColumnIndex < columns.firstColumnIndex()) {
            shifted.put(existingColumnIndex, definition);
            return;
          }
          if (existingColumnIndex > columns.lastColumnIndex()) {
            shifted.put(existingColumnIndex - columns.count(), definition);
          }
        });
    return Map.copyOf(shifted);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  public static Map<Integer, CTCol> shiftedForShift(
      Map<Integer, CTCol> explicitColumns, ExcelColumnSpan columns, int delta) {
    int destinationFirstColumn = columns.firstColumnIndex() + delta;
    int destinationLastColumn = columns.lastColumnIndex() + delta;
    int overwrittenFirstColumn = Math.min(destinationFirstColumn, destinationLastColumn);
    int overwrittenLastColumn = Math.max(destinationFirstColumn, destinationLastColumn);
    Map<Integer, CTCol> shifted = new LinkedHashMap<>();
    explicitColumns.forEach(
        (existingColumnIndex, definition) -> {
          boolean inSource =
              existingColumnIndex >= columns.firstColumnIndex()
                  && existingColumnIndex <= columns.lastColumnIndex();
          boolean overwrittenDestination =
              existingColumnIndex >= overwrittenFirstColumn
                  && existingColumnIndex <= overwrittenLastColumn;
          if (!inSource && !overwrittenDestination) {
            shifted.put(existingColumnIndex, definition);
          }
        });
    explicitColumns.forEach(
        (existingColumnIndex, definition) -> {
          if (existingColumnIndex >= columns.firstColumnIndex()
              && existingColumnIndex <= columns.lastColumnIndex()) {
            shifted.put(existingColumnIndex + delta, definition);
          }
        });
    return Map.copyOf(shifted);
  }

  private static void rebuildColumnDefinitionsInternal(
      XSSFSheet sheet, Map<Integer, CTCol> explicitColumns) {
    CTCols cols = CTCols.Factory.newInstance();
    explicitColumns.entrySet().stream()
        .sorted(Map.Entry.comparingByKey())
        .forEach(
            entry -> {
              CTCol definition = copyOf(entry.getValue());
              definition.setMin(entry.getKey() + 1L);
              definition.setMax(entry.getKey() + 1L);
              cols.addNewCol().set(definition);
            });
    sheet.getCTWorksheet().setColsArray(new CTCols[] {cols});
  }

  private static boolean requiresColumnDefinitionCanonicalization(XSSFSheet sheet) {
    if (sheet.getCTWorksheet().sizeOfColsArray() != 1) {
      return true;
    }
    boolean[] seenColumns = new boolean[ExcelColumnSpan.MAX_COLUMN_INDEX + 1];
    for (CTCol col : sheet.getCTWorksheet().getColsArray(0).getColList()) {
      if (col.getMin() != col.getMax()) {
        return true;
      }
      if (isSemanticallyEmptyColumnDefinition(col)) {
        return true;
      }
      for (int columnIndex = (int) col.getMin() - 1;
          columnIndex <= (int) col.getMax() - 1;
          columnIndex++) {
        if (seenColumns[columnIndex]) {
          return true;
        }
        seenColumns[columnIndex] = true;
      }
    }
    return false;
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<Integer, CTCol> effectiveColumnDefinitions(XSSFSheet sheet) {
    Map<Integer, CTCol> explicitColumns = snapshotColumnDefinitions(sheet);
    int lastColumnIndex = lastColumnIndex(sheet);
    if (lastColumnIndex < 0) {
      return Map.of();
    }
    Map<Integer, CTCol> effectiveColumns = new LinkedHashMap<>();
    for (int columnIndex = 0; columnIndex <= lastColumnIndex; columnIndex++) {
      int targetColumnIndex = columnIndex;
      effectiveColumnDefinition(sheet, columnIndex, explicitColumns.get(columnIndex))
          .ifPresent(
              effectiveDefinition -> effectiveColumns.put(targetColumnIndex, effectiveDefinition));
    }
    return Map.copyOf(effectiveColumns);
  }

  private static Optional<CTCol> effectiveColumnDefinition(
      XSSFSheet sheet, int columnIndex, @Nullable CTCol baseDefinition) {
    if (baseDefinition == null) {
      return Optional.empty();
    }
    boolean hidden = false;
    long outlineLevel = 0L;
    boolean collapsed = false;
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        if (columnIndex + 1 < col.getMin() || columnIndex + 1 > col.getMax()) {
          continue;
        }
        hidden |= col.getHidden();
        outlineLevel = Math.max(outlineLevel, col.getOutlineLevel());
        collapsed |= col.getCollapsed();
      }
    }
    if (!hasMeaningfulColumnDefinition(baseDefinition, hidden, outlineLevel, collapsed)) {
      return Optional.empty();
    }
    CTCol effectiveDefinition = copyOf(baseDefinition);
    effectiveDefinition.setMin(columnIndex + 1L);
    effectiveDefinition.setMax(columnIndex + 1L);
    effectiveDefinition.setHidden(hidden);
    effectiveDefinition.setOutlineLevel((short) outlineLevel);
    effectiveDefinition.setCollapsed(collapsed);
    return Optional.of(effectiveDefinition);
  }

  private static boolean hasMeaningfulColumnDefinition(
      CTCol baseDefinition, boolean hidden, long outlineLevel, boolean collapsed) {
    return hidden
        || outlineLevel > 0L
        || collapsed
        || !isSemanticallyEmptyColumnDefinition(baseDefinition);
  }

  private static boolean isSemanticallyEmptyColumnDefinition(CTCol definition) {
    return !definition.getHidden()
        && definition.getOutlineLevel() == 0
        && !definition.getCollapsed()
        && !definition.getCustomWidth()
        && !definition.getBestFit()
        && !definition.getPhonetic()
        && (!definition.isSetStyle() || definition.getStyle() == 0L);
  }

  private static CTCol copyOf(CTCol original) {
    Objects.requireNonNull(original, "original must not be null");
    CTCol copy = CTCol.Factory.newInstance();
    copy.set(original);
    return copy;
  }
}
