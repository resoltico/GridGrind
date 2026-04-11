package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;

/** Reads, writes, and analyzes sheet-owned autofilter structures on one XSSF sheet. */
final class ExcelAutofilterController {
  /** Creates or replaces one sheet-level autofilter range. */
  void setSheetAutofilter(XSSFSheet sheet, String range) {
    setSheetAutofilter(sheet, range, List.of(), null);
  }

  /** Creates or replaces one sheet-level autofilter range. */
  void setSheetAutofilter(
      XSSFSheet sheet,
      String range,
      List<ExcelAutofilterFilterColumn> criteria,
      ExcelAutofilterSortState sortState) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(range, "range must not be null");
    List<ExcelAutofilterFilterColumn> filterCriteria =
        List.copyOf(Objects.requireNonNull(criteria, "criteria must not be null"));

    ExcelRange targetRange = ExcelRange.parse(range);
    if (ExcelSheetStructureSupport.headerRowMissing(sheet, targetRange)) {
      throw new IllegalArgumentException(
          "autofilter range must include a nonblank header row: "
              + ExcelSheetStructureSupport.formatRange(targetRange));
    }
    requireNoTableOverlap(sheet, targetRange);
    sheet.setAutoFilter(ExcelSheetStructureSupport.toCellRangeAddress(targetRange));
    CTAutoFilter autoFilter = sheet.getCTWorksheet().getAutoFilter();
    replaceFilterColumns(sheet.getWorkbook(), autoFilter, filterCriteria);
    replaceSortState(sheet.getWorkbook(), autoFilter, sortState);
  }

  /** Clears the sheet-level autofilter range on one sheet. */
  void clearSheetAutofilter(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (sheet.getCTWorksheet().isSetAutoFilter()) {
      sheet.getCTWorksheet().unsetAutoFilter();
    }
    List<Name> filterDatabaseNames = new ArrayList<>();
    int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
    for (Name name : sheet.getWorkbook().getAllNames()) {
      if (name.getSheetIndex() == sheetIndex
          && "_XLNM._FILTERDATABASE".equalsIgnoreCase(name.getNameName())) {
        filterDatabaseNames.add(name);
      }
    }
    for (Name filterDatabaseName : filterDatabaseNames) {
      sheet.getWorkbook().removeName(filterDatabaseName);
    }
  }

  /** Returns the sheet-owned autofilter metadata present on one sheet. */
  List<ExcelAutofilterSnapshot> sheetOwnedAutofilters(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (!sheet.getCTWorksheet().isSetAutoFilter()) {
      return List.of();
    }
    CTAutoFilter autoFilter = sheet.getCTWorksheet().getAutoFilter();
    return List.of(
        new ExcelAutofilterSnapshot.SheetOwned(
            Objects.requireNonNullElse(autoFilter.getRef(), ""),
            filterColumns(sheet.getWorkbook(), autoFilter),
            sortState(sheet.getWorkbook(), autoFilter)));
  }

  /** Returns the number of sheet-owned autofilters currently present on one sheet. */
  int sheetAutofilterCount(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.getCTWorksheet().isSetAutoFilter() ? 1 : 0;
  }

  /** Returns derived health findings for the sheet-owned autofilter on one sheet. */
  List<WorkbookAnalysis.AnalysisFinding> sheetAutofilterHealthFindings(
      String sheetName, XSSFSheet sheet, List<ExcelTableSnapshot> tablesOnSheet) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(tablesOnSheet, "tablesOnSheet must not be null");

    if (!sheet.getCTWorksheet().isSetAutoFilter()) {
      return List.of();
    }

    String rawRange =
        Objects.requireNonNullElse(sheet.getCTWorksheet().getAutoFilter().getRef(), "");
    ExcelRange targetRange = ExcelSheetStructureSupport.parseRangeOrNull(rawRange);
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    if (targetRange == null) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_INVALID_RANGE,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              "Autofilter range is invalid",
              "Sheet-owned autofilter range could not be parsed.",
              new WorkbookAnalysis.AnalysisLocation.Sheet(sheetName),
              List.of(rawRange)));
      return List.copyOf(findings);
    }

    String normalizedRange = ExcelSheetStructureSupport.formatRange(targetRange);
    WorkbookAnalysis.AnalysisLocation.Range location =
        new WorkbookAnalysis.AnalysisLocation.Range(sheetName, normalizedRange);
    if (ExcelSheetStructureSupport.headerRowMissing(sheet, targetRange)) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_MISSING_HEADER_ROW,
              WorkbookAnalysis.AnalysisSeverity.WARNING,
              "Autofilter header row is blank",
              "Sheet-owned autofilter range does not contain a nonblank header row.",
              location,
              List.of(normalizedRange)));
    }

    List<String> overlappingTables = new ArrayList<>();
    for (ExcelTableSnapshot table : tablesOnSheet) {
      ExcelRange tableRange = ExcelSheetStructureSupport.parseRangeOrNull(table.range());
      if (tableRange != null && ExcelSheetStructureSupport.intersects(targetRange, tableRange)) {
        overlappingTables.add(table.name() + "@" + table.range());
      }
    }
    if (!overlappingTables.isEmpty()) {
      List<String> evidence = new ArrayList<>();
      evidence.add(normalizedRange);
      evidence.addAll(overlappingTables);
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
              WorkbookAnalysis.AnalysisSeverity.WARNING,
              "Sheet autofilter overlaps a table range",
              "Sheet-owned autofilter metadata overlaps one or more table ranges."
                  + " Table-owned filters should be managed by table definitions instead.",
              location,
              List.copyOf(evidence)));
    }
    return List.copyOf(findings);
  }

  private void requireNoTableOverlap(XSSFSheet sheet, ExcelRange targetRange) {
    List<String> overlaps = new ArrayList<>();
    for (var table : sheet.getTables()) {
      ExcelRange tableRange =
          ExcelSheetStructureSupport.parseRangeOrNull(
              Objects.requireNonNullElse(table.getCTTable().getRef(), ""));
      if (tableRange != null && ExcelSheetStructureSupport.intersects(targetRange, tableRange)) {
        overlaps.add(table.getName() + "@" + table.getCTTable().getRef());
      }
    }
    if (!overlaps.isEmpty()) {
      throw new IllegalArgumentException(
          "sheet-level autofilter range must not overlap an existing table range: "
              + String.join(", ", overlaps));
    }
  }

  private static void replaceFilterColumns(
      XSSFWorkbook workbook, CTAutoFilter autoFilter, List<ExcelAutofilterFilterColumn> criteria) {
    while (autoFilter.sizeOfFilterColumnArray() > 0) {
      autoFilter.removeFilterColumn(0);
    }
    for (ExcelAutofilterFilterColumn column : criteria) {
      CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
      filterColumn.setColId(column.columnId());
      if (!column.showButton()) {
        filterColumn.setShowButton(false);
      }
      applyCriterion(workbook, filterColumn, column.criterion());
    }
  }

  private static void applyCriterion(
      XSSFWorkbook workbook,
      CTFilterColumn filterColumn,
      ExcelAutofilterFilterCriterion criterion) {
    switch (criterion) {
      case ExcelAutofilterFilterCriterion.Values values -> {
        CTFilters filters = filterColumn.addNewFilters();
        for (String value : values.values()) {
          filters.addNewFilter().setVal(value);
        }
        if (values.includeBlank()) {
          filters.setBlank(true);
        }
      }
      case ExcelAutofilterFilterCriterion.Custom custom -> {
        CTCustomFilters customFilters = filterColumn.addNewCustomFilters();
        customFilters.setAnd(custom.and());
        for (ExcelAutofilterFilterCriterion.CustomCondition condition : custom.conditions()) {
          CTCustomFilter customFilter = customFilters.addNewCustomFilter();
          STFilterOperator.Enum operator = STFilterOperator.Enum.forString(condition.operator());
          if (operator == null) {
            throw new IllegalArgumentException(
                "unsupported autofilter custom operator: " + condition.operator());
          }
          customFilter.setOperator(operator);
          customFilter.setVal(condition.value());
        }
      }
      case ExcelAutofilterFilterCriterion.Dynamic dynamic -> {
        CTDynamicFilter dynamicFilter = filterColumn.addNewDynamicFilter();
        STDynamicFilterType.Enum type = STDynamicFilterType.Enum.forString(dynamic.type());
        if (type == null) {
          throw new IllegalArgumentException(
              "unsupported autofilter dynamic type: " + dynamic.type());
        }
        dynamicFilter.setType(type);
        if (dynamic.value() != null) {
          dynamicFilter.setVal(dynamic.value());
        }
        if (dynamic.maxValue() != null) {
          dynamicFilter.setMaxVal(dynamic.maxValue());
        }
      }
      case ExcelAutofilterFilterCriterion.Top10 top10 -> {
        CTTop10 top10Filter = filterColumn.addNewTop10();
        top10Filter.setVal(top10.value());
        top10Filter.setTop(top10.top());
        top10Filter.setPercent(top10.percent());
      }
      case ExcelAutofilterFilterCriterion.Color color -> {
        CTColorFilter colorFilter = filterColumn.addNewColorFilter();
        colorFilter.setCellColor(color.cellColor());
        colorFilter.setDxfId(putColorDxf(workbook, color.color(), color.cellColor()) - 1L);
      }
      case ExcelAutofilterFilterCriterion.Icon icon -> {
        CTIconFilter iconFilter = filterColumn.addNewIconFilter();
        STIconSetType.Enum iconSet = STIconSetType.Enum.forString(icon.iconSet());
        if (iconSet == null) {
          throw new IllegalArgumentException("unsupported autofilter icon set: " + icon.iconSet());
        }
        iconFilter.setIconSet(iconSet);
        iconFilter.setIconId(icon.iconId());
      }
    }
  }

  private static void replaceSortState(
      XSSFWorkbook workbook, CTAutoFilter autoFilter, ExcelAutofilterSortState sortState) {
    if (autoFilter.isSetSortState()) {
      autoFilter.unsetSortState();
    }
    if (sortState == null) {
      return;
    }
    CTSortState ctSortState = autoFilter.addNewSortState();
    applySortStateSettings(ctSortState, sortState);
    for (ExcelAutofilterSortCondition condition : sortState.conditions()) {
      applySortCondition(workbook, ctSortState, condition);
    }
  }

  private static void applySortStateSettings(
      CTSortState ctSortState, ExcelAutofilterSortState sortState) {
    ctSortState.setRef(sortState.range());
    if (sortState.caseSensitive()) {
      ctSortState.setCaseSensitive(true);
    }
    if (sortState.columnSort()) {
      ctSortState.setColumnSort(true);
    }
    if (!sortState.sortMethod().isBlank()) {
      STSortMethod.Enum sortMethod = STSortMethod.Enum.forString(sortState.sortMethod());
      if (sortMethod == null) {
        throw new IllegalArgumentException(
            "unsupported autofilter sort method: " + sortState.sortMethod());
      }
      ctSortState.setSortMethod(sortMethod);
    }
  }

  private static void applySortCondition(
      XSSFWorkbook workbook, CTSortState ctSortState, ExcelAutofilterSortCondition condition) {
    CTSortCondition sortCondition = ctSortState.addNewSortCondition();
    sortCondition.setRef(condition.range());
    if (condition.descending()) {
      sortCondition.setDescending(true);
    }
    if (!condition.sortBy().isBlank()) {
      STSortBy.Enum sortBy = STSortBy.Enum.forString(condition.sortBy());
      if (sortBy == null) {
        throw new IllegalArgumentException(
            "unsupported autofilter sortBy value: " + condition.sortBy());
      }
      sortCondition.setSortBy(sortBy);
    }
    if (condition.color() != null) {
      sortCondition.setDxfId(putColorDxf(workbook, condition.color(), true) - 1L);
    }
    if (condition.iconId() != null) {
      sortCondition.setIconId(condition.iconId());
    }
  }

  private static long putColorDxf(XSSFWorkbook workbook, ExcelColor color, boolean cellColor) {
    CTDxf dxf = CTDxf.Factory.newInstance();
    if (cellColor) {
      CTPatternFill patternFill = dxf.addNewFill().addNewPatternFill();
      patternFill.setPatternType(STPatternType.SOLID);
      patternFill.addNewFgColor().set(ExcelColorSupport.toXssfColor(workbook, color).getCTColor());
    } else {
      CTFont font = dxf.addNewFont();
      font.addNewColor().set(ExcelColorSupport.toXssfColor(workbook, color).getCTColor());
    }
    return workbook.getStylesSource().putDxf(dxf);
  }

  List<ExcelAutofilterFilterColumnSnapshot> filterColumns(
      XSSFWorkbook workbook, CTAutoFilter autoFilter) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(autoFilter, "autoFilter must not be null");

    return Arrays.stream(autoFilter.getFilterColumnArray())
        .map(filterColumn -> filterColumnSnapshot(workbook, filterColumn))
        .toList();
  }

  ExcelAutofilterSortStateSnapshot sortState(XSSFWorkbook workbook, CTAutoFilter autoFilter) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(autoFilter, "autoFilter must not be null");
    if (!autoFilter.isSetSortState()) {
      return null;
    }
    CTSortState sortState = autoFilter.getSortState();
    return new ExcelAutofilterSortStateSnapshot(
        Objects.requireNonNullElse(sortState.getRef(), ""),
        sortState.isSetCaseSensitive() && sortState.getCaseSensitive(),
        sortState.isSetColumnSort() && sortState.getColumnSort(),
        sortState.isSetSortMethod() ? sortState.getSortMethod().toString() : "",
        Arrays.stream(sortState.getSortConditionArray())
            .map(condition -> sortConditionSnapshot(workbook, condition))
            .toList());
  }

  private static ExcelAutofilterFilterColumnSnapshot filterColumnSnapshot(
      XSSFWorkbook workbook, CTFilterColumn filterColumn) {
    return new ExcelAutofilterFilterColumnSnapshot(
        filterColumn.getColId(),
        showButton(filterColumn),
        criterionSnapshot(workbook, filterColumn));
  }

  private static boolean showButton(CTFilterColumn filterColumn) {
    if (filterColumn.isSetShowButton()) {
      return filterColumn.getShowButton();
    }
    return !filterColumn.isSetHiddenButton() || !filterColumn.getHiddenButton();
  }

  private static ExcelAutofilterFilterCriterionSnapshot criterionSnapshot(
      XSSFWorkbook workbook, CTFilterColumn filterColumn) {
    if (filterColumn.isSetFilters()) {
      return valuesCriterion(filterColumn.getFilters());
    }
    if (filterColumn.isSetCustomFilters()) {
      return customCriterion(filterColumn.getCustomFilters());
    }
    if (filterColumn.isSetDynamicFilter()) {
      return dynamicCriterion(filterColumn.getDynamicFilter());
    }
    if (filterColumn.isSetTop10()) {
      return top10Criterion(filterColumn.getTop10());
    }
    if (filterColumn.isSetColorFilter()) {
      return colorCriterion(workbook, filterColumn.getColorFilter());
    }
    if (filterColumn.isSetIconFilter()) {
      return iconCriterion(filterColumn.getIconFilter());
    }
    return new ExcelAutofilterFilterCriterionSnapshot.Values(List.of(), false);
  }

  static ExcelAutofilterFilterCriterionSnapshot valuesCriterion(CTFilters filters) {
    List<String> values = new ArrayList<>();
    for (CTFilter filter : filters.getFilterArray()) {
      values.add(Objects.requireNonNullElse(filter.getVal(), ""));
    }
    return new ExcelAutofilterFilterCriterionSnapshot.Values(
        List.copyOf(values), filters.isSetBlank() && filters.getBlank());
  }

  static ExcelAutofilterFilterCriterionSnapshot customCriterion(CTCustomFilters customFilters) {
    return new ExcelAutofilterFilterCriterionSnapshot.Custom(
        customFilters.isSetAnd() && customFilters.getAnd(),
        Arrays.stream(customFilters.getCustomFilterArray())
            .map(ExcelAutofilterController::customCondition)
            .toList());
  }

  private static ExcelAutofilterFilterCriterionSnapshot.CustomCondition customCondition(
      CTCustomFilter customFilter) {
    String operator =
        customFilter.isSetOperator() ? customFilter.getOperator().toString() : "equal";
    return new ExcelAutofilterFilterCriterionSnapshot.CustomCondition(
        operator, Objects.requireNonNullElse(customFilter.getVal(), ""));
  }

  static ExcelAutofilterFilterCriterionSnapshot dynamicCriterion(CTDynamicFilter dynamicFilter) {
    return new ExcelAutofilterFilterCriterionSnapshot.Dynamic(
        dynamicFilter.getType() == null ? "UNKNOWN" : dynamicFilter.getType().toString(),
        dynamicFilter.isSetVal() ? dynamicFilter.getVal() : null,
        dynamicFilter.isSetMaxVal() ? dynamicFilter.getMaxVal() : null);
  }

  static ExcelAutofilterFilterCriterionSnapshot top10Criterion(CTTop10 top10) {
    return new ExcelAutofilterFilterCriterionSnapshot.Top10(
        !top10.isSetTop() || top10.getTop(),
        top10.isSetPercent() && top10.getPercent(),
        top10.getVal(),
        top10.isSetFilterVal() ? top10.getFilterVal() : null);
  }

  static ExcelAutofilterFilterCriterionSnapshot colorCriterion(
      XSSFWorkbook workbook, CTColorFilter colorFilter) {
    return new ExcelAutofilterFilterCriterionSnapshot.Color(
        colorFilter.isSetCellColor() && colorFilter.getCellColor(),
        colorFilter.isSetDxfId()
            ? dxfColor(
                workbook,
                colorFilter.getDxfId(),
                colorFilter.isSetCellColor() && colorFilter.getCellColor())
            : null);
  }

  static ExcelAutofilterFilterCriterionSnapshot iconCriterion(CTIconFilter iconFilter) {
    return new ExcelAutofilterFilterCriterionSnapshot.Icon(
        iconFilter.getIconSet() == null ? "UNKNOWN" : iconFilter.getIconSet().toString(),
        iconFilter.isSetIconId() ? Math.toIntExact(iconFilter.getIconId()) : 0);
  }

  static ExcelAutofilterSortConditionSnapshot sortConditionSnapshot(
      XSSFWorkbook workbook, CTSortCondition sortCondition) {
    boolean sortByCellColor =
        sortCondition.isSetSortBy()
            && "cellColor".equalsIgnoreCase(sortCondition.getSortBy().toString());
    boolean sortByFontColor =
        sortCondition.isSetSortBy()
            && "fontColor".equalsIgnoreCase(sortCondition.getSortBy().toString());
    return new ExcelAutofilterSortConditionSnapshot(
        Objects.requireNonNullElse(sortCondition.getRef(), ""),
        sortCondition.isSetDescending() && sortCondition.getDescending(),
        sortCondition.isSetSortBy() ? sortCondition.getSortBy().toString() : "",
        sortCondition.isSetDxfId()
            ? dxfColor(workbook, sortCondition.getDxfId(), sortByCellColor || !sortByFontColor)
            : null,
        sortCondition.isSetIconId() ? Math.toIntExact(sortCondition.getIconId()) : null);
  }

  static ExcelColorSnapshot dxfColor(XSSFWorkbook workbook, long dxfId, boolean cellColor) {
    CTDxf dxf = dxfAt(workbook.getStylesSource(), dxfId);
    if (dxf == null) {
      return null;
    }
    if (cellColor
        && dxf.isSetFill()
        && dxf.getFill().isSetPatternFill()
        && dxf.getFill().getPatternFill().isSetFgColor()) {
      return ExcelColorSnapshotSupport.snapshot(
          workbook, dxf.getFill().getPatternFill().getFgColor());
    }
    if (!cellColor && dxf.isSetFont() && dxf.getFont().sizeOfColorArray() > 0) {
      return ExcelColorSnapshotSupport.snapshot(workbook, dxf.getFont().getColorArray(0));
    }
    if (dxf.isSetFill()
        && dxf.getFill().isSetPatternFill()
        && dxf.getFill().getPatternFill().isSetFgColor()) {
      return ExcelColorSnapshotSupport.snapshot(
          workbook, dxf.getFill().getPatternFill().getFgColor());
    }
    if (dxf.isSetFont() && dxf.getFont().sizeOfColorArray() > 0) {
      return ExcelColorSnapshotSupport.snapshot(workbook, dxf.getFont().getColorArray(0));
    }
    return null;
  }

  static CTDxf dxfAt(StylesTable stylesTable, long dxfId) {
    if (dxfId < 0L || dxfId >= stylesTable._getDXfsSize()) {
      return null;
    }
    return stylesTable.getDxfAt(Math.toIntExact(dxfId));
  }
}
