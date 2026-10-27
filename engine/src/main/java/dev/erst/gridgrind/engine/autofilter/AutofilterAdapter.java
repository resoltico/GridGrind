package dev.erst.gridgrind.engine.autofilter;

import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.vocabulary.ExcelAutofilterSortMethod;
import dev.erst.gridgrind.engine.poi.CellBounds;
import dev.erst.gridgrind.engine.poi.SheetStructure;
import dev.erst.gridgrind.engine.styles.ColorAdapter;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColorFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCustomFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCustomFilters;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDynamicFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilterColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilters;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTIconFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSortCondition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSortState;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTop10;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDynamicFilterType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STFilterOperator;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STIconSetType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSortBy;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSortMethod;

/** Owns sheet/table autofilter criteria, sorting, differential colors and readback. */
@SuppressWarnings(
    "PMD.ExcessiveImports") // One cohesive adapter owns the complete SpreadsheetML autofilter
// surface.
public final class AutofilterAdapter {
  /** Creates or replaces one sheet-level autofilter range. */
  public void setSheetAutofilter(XSSFSheet sheet, String range) {
    setSheetAutofilter(sheet, range, List.of(), Optional.empty());
  }

  /** Creates or replaces one sheet-level autofilter range. */
  public void setSheetAutofilter(
      XSSFSheet sheet,
      String range,
      List<AutofilterFilterColumnSpec> criteria,
      Optional<AutofilterSortStateSpec> sortState) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(range, "range must not be null");
    List<AutofilterFilterColumnSpec> filterCriteria =
        List.copyOf(Objects.requireNonNull(criteria, "criteria must not be null"));
    Objects.requireNonNull(sortState, "sortState must not be null");

    CellBounds targetRange = CellBounds.parse(range);
    if (SheetStructure.headerRowMissing(sheet, targetRange)) {
      throw new IllegalArgumentException(
          "autofilter range must include a nonblank header row: "
              + SheetStructure.formatRange(targetRange));
    }
    requireNoTableOverlap(sheet, targetRange);
    sheet.setAutoFilter(SheetStructure.toCellRangeAddress(targetRange));
    CTAutoFilter autoFilter = sheet.getCTWorksheet().getAutoFilter();
    replaceFilterColumns(sheet.getWorkbook(), autoFilter, filterCriteria);
    replaceSortState(sheet.getWorkbook(), autoFilter, sortState);
  }

  /** Clears the sheet-level autofilter range on one sheet. */
  public void clearSheetAutofilter(XSSFSheet sheet) {
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
  public List<AutofilterEntryReport> sheetOwnedAutofilters(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (!sheet.getCTWorksheet().isSetAutoFilter()) {
      return List.of();
    }
    CTAutoFilter autoFilter = sheet.getCTWorksheet().getAutoFilter();
    return List.of(
        new AutofilterEntryReport.SheetOwned(
            new dev.erst.gridgrind.core.vocabulary.RangeLabel(
                Objects.requireNonNullElse(autoFilter.getRef(), "")),
            filterColumns(sheet.getWorkbook(), autoFilter),
            sortState(sheet.getWorkbook(), autoFilter)));
  }

  /** Returns the number of sheet-owned autofilters currently present on one sheet. */
  public int sheetAutofilterCount(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.getCTWorksheet().isSetAutoFilter() ? 1 : 0;
  }

  /** Returns derived health findings for the sheet-owned autofilter on one sheet. */
  public List<AnalysisFindingReport> sheetAutofilterHealthFindings(
      String sheetName, XSSFSheet sheet, List<TableEntryReport> tablesOnSheet) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(tablesOnSheet, "tablesOnSheet must not be null");

    if (!sheet.getCTWorksheet().isSetAutoFilter()) {
      return List.of();
    }

    String rawRange =
        Objects.requireNonNullElse(sheet.getCTWorksheet().getAutoFilter().getRef(), "");
    Optional<CellBounds> targetRange = SheetStructure.parseOptionalRange(rawRange);
    List<AnalysisFindingReport> findings = new ArrayList<>();
    if (targetRange.isEmpty()) {
      findings.add(
          new AnalysisFindingReport(
              dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode.AUTOFILTER_INVALID_RANGE,
              dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.ERROR,
              "Autofilter range is invalid",
              "Sheet-owned autofilter range could not be parsed.",
              new AnalysisLocationReport.Sheet(
                  new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName)),
              List.of(rawRange)));
      return List.copyOf(findings);
    }

    CellBounds parsedTargetRange = targetRange.orElseThrow();
    String normalizedRange = SheetStructure.formatRange(parsedTargetRange);
    AnalysisLocationReport.Range location =
        new AnalysisLocationReport.Range(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName), normalizedRange);
    if (SheetStructure.headerRowMissing(sheet, parsedTargetRange)) {
      findings.add(
          new AnalysisFindingReport(
              dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode.AUTOFILTER_MISSING_HEADER_ROW,
              dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.WARNING,
              "Autofilter header row is blank",
              "Sheet-owned autofilter range does not contain a nonblank header row.",
              location,
              List.of(normalizedRange)));
    }

    List<String> overlappingTables = new ArrayList<>();
    for (TableEntryReport table : tablesOnSheet) {
      Optional<CellBounds> tableRange = SheetStructure.parseOptionalRange(table.range().value());
      if (tableRange.isPresent()
          && SheetStructure.intersects(parsedTargetRange, tableRange.orElseThrow())) {
        overlappingTables.add(table.name() + "@" + table.range());
      }
    }
    if (!overlappingTables.isEmpty()) {
      List<String> evidence = new ArrayList<>();
      evidence.add(normalizedRange);
      evidence.addAll(overlappingTables);
      findings.add(
          new AnalysisFindingReport(
              dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
              dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.WARNING,
              "Sheet autofilter overlaps a table range",
              "Sheet-owned autofilter metadata overlaps one or more table ranges."
                  + " Table-owned filters should be managed by table definitions instead.",
              location,
              List.copyOf(evidence)));
    }
    return List.copyOf(findings);
  }

  private void requireNoTableOverlap(XSSFSheet sheet, CellBounds targetRange) {
    List<String> overlaps = new ArrayList<>();
    for (var table : sheet.getTables()) {
      Optional<CellBounds> tableRange =
          SheetStructure.parseOptionalRange(
              Objects.requireNonNullElse(table.getCTTable().getRef(), ""));
      if (tableRange.isPresent()
          && SheetStructure.intersects(targetRange, tableRange.orElseThrow())) {
        overlaps.add(table.getName() + "@" + table.getCTTable().getRef());
      }
    }
    if (!overlaps.isEmpty()) {
      throw new IllegalArgumentException(
          "sheet-level autofilter range must not overlap an existing table range: "
              + String.join(", ", overlaps));
    }
  }

  /** Replaces authored filter columns while preserving their explicit criteria. */
  public static void replaceFilterColumns(
      XSSFWorkbook workbook, CTAutoFilter autoFilter, List<AutofilterFilterColumnSpec> criteria) {
    while (autoFilter.sizeOfFilterColumnArray() > 0) {
      autoFilter.removeFilterColumn(0);
    }
    for (AutofilterFilterColumnSpec column : criteria) {
      CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
      filterColumn.setColId(column.columnId());
      if (!column.showButton()) {
        filterColumn.setShowButton(false);
      }
      applyCriterion(workbook, filterColumn, column.criterion());
    }
  }

  /** Replaces the optional sheet or table autofilter sort state. */
  public static void replaceSortState(
      XSSFWorkbook workbook, CTAutoFilter autoFilter, Optional<AutofilterSortStateSpec> sortState) {
    if (autoFilter.isSetSortState()) {
      autoFilter.unsetSortState();
    }
    if (sortState.isEmpty()) {
      return;
    }
    CTSortState ctSortState = autoFilter.addNewSortState();
    AutofilterSortStateSpec authoredSortState = sortState.orElseThrow();
    applySortStateSettings(ctSortState, authoredSortState);
    for (AutofilterSortConditionSpec condition : authoredSortState.conditions()) {
      applySortCondition(workbook, ctSortState, condition);
    }
  }

  private static void applySortStateSettings(
      CTSortState ctSortState, AutofilterSortStateSpec sortState) {
    ctSortState.setRef(sortState.range().value());
    if (sortState.caseSensitive()) {
      ctSortState.setCaseSensitive(true);
    }
    if (sortState.columnSort()) {
      ctSortState.setColumnSort(true);
    }
    sortState
        .sortMethod()
        .ifPresent(sortMethod -> ctSortState.setSortMethod(toOoxmlSortMethod(sortMethod)));
  }

  private static void applySortCondition(
      XSSFWorkbook workbook, CTSortState ctSortState, AutofilterSortConditionSpec condition) {
    CTSortCondition sortCondition = ctSortState.addNewSortCondition();
    sortCondition.setRef(condition.range().value());
    if (condition.descending()) {
      sortCondition.setDescending(true);
    }
    switch (condition) {
      case AutofilterSortConditionSpec.Value _ -> {
        // SpreadsheetML uses VALUE semantics when no explicit sortBy discriminator is present.
      }
      case AutofilterSortConditionSpec.CellColor cellColor -> {
        sortCondition.setSortBy(STSortBy.CELL_COLOR);
        sortCondition.setDxfId(putColorDxf(workbook, cellColor.color(), true) - 1L);
      }
      case AutofilterSortConditionSpec.FontColor fontColor -> {
        sortCondition.setSortBy(STSortBy.FONT_COLOR);
        sortCondition.setDxfId(putColorDxf(workbook, fontColor.color(), false) - 1L);
      }
      case AutofilterSortConditionSpec.Icon icon -> {
        sortCondition.setSortBy(STSortBy.ICON);
        sortCondition.setIconId(icon.iconId());
      }
    }
  }

  /** Reads effective filter-column criteria from native XML. */
  public static List<AutofilterFilterColumnReport> filterColumns(
      XSSFWorkbook workbook, CTAutoFilter autoFilter) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(autoFilter, "autoFilter must not be null");
    return java.util.Arrays.stream(autoFilter.getFilterColumnArray())
        .map(filterColumn -> filterColumnSnapshot(workbook, filterColumn))
        .toList();
  }

  /** Reads the optional native sort state and its conditions. */
  public static Optional<AutofilterSortStateReport> sortState(
      XSSFWorkbook workbook, CTAutoFilter autoFilter) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(autoFilter, "autoFilter must not be null");
    if (!autoFilter.isSetSortState()) {
      return Optional.empty();
    }
    CTSortState sortState = autoFilter.getSortState();
    String range =
        sortState.getRef() != null
            ? sortState.getRef()
            : Objects.requireNonNullElse(autoFilter.getRef(), "");
    if (range.isBlank()) {
      throw new IllegalArgumentException("autofilter sort state is missing ref");
    }
    return Optional.of(
        new AutofilterSortStateReport(
            new dev.erst.gridgrind.core.vocabulary.RangeLabel(range),
            sortState.isSetCaseSensitive() && sortState.getCaseSensitive(),
            sortState.isSetColumnSort() && sortState.getColumnSort(),
            sortState.isSetSortMethod()
                ? ExcelAutofilterSortMethod.fromOoxmlValue(sortState.getSortMethod().toString())
                : Optional.empty(),
            java.util.Arrays.stream(sortState.getSortConditionArray())
                .map(condition -> sortConditionSnapshot(workbook, condition))
                .toList()));
  }

  private static AutofilterFilterColumnReport filterColumnSnapshot(
      XSSFWorkbook workbook, CTFilterColumn filterColumn) {
    return new AutofilterFilterColumnReport(
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

  private static AutofilterFilterCriterionReport criterionSnapshot(
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
    return new AutofilterFilterCriterionReport.Values(List.of(), false);
  }

  /** Reads explicit values and blank inclusion from a value filter. */
  public static AutofilterFilterCriterionReport valuesCriterion(CTFilters filters) {
    List<String> values = new ArrayList<>();
    for (CTFilter filter : filters.getFilterArray()) {
      values.add(Objects.requireNonNullElse(filter.getVal(), ""));
    }
    return new AutofilterFilterCriterionReport.Values(
        List.copyOf(values), filters.isSetBlank() && filters.getBlank());
  }

  /** Reads custom comparison conditions and their conjunction. */
  public static AutofilterFilterCriterionReport customCriterion(CTCustomFilters customFilters) {
    return new AutofilterFilterCriterionReport.Custom(
        customFilters.isSetAnd() && customFilters.getAnd(),
        java.util.Arrays.stream(customFilters.getCustomFilterArray())
            .map(AutofilterAdapter::customCondition)
            .toList());
  }

  private static AutofilterFilterCriterionReport.CustomConditionReport customCondition(
      CTCustomFilter customFilter) {
    String operator =
        customFilter.isSetOperator() ? customFilter.getOperator().toString() : "equal";
    return new AutofilterFilterCriterionReport.CustomConditionReport(
        operator, Objects.requireNonNullElse(customFilter.getVal(), ""));
  }

  /** Reads the dynamic filter kind and numeric bounds. */
  public static AutofilterFilterCriterionReport dynamicCriterion(CTDynamicFilter dynamicFilter) {
    return new AutofilterFilterCriterionReport.Dynamic(
        dynamicFilter.getType() == null ? "UNKNOWN" : dynamicFilter.getType().toString(),
        Optional.ofNullable(dynamicFilter.isSetVal() ? dynamicFilter.getVal() : null),
        Optional.ofNullable(dynamicFilter.isSetMaxVal() ? dynamicFilter.getMaxVal() : null));
  }

  /** Reads the native top/bottom count or percent criterion. */
  public static AutofilterFilterCriterionReport top10Criterion(CTTop10 top10) {
    return new AutofilterFilterCriterionReport.Top10(
        !top10.isSetTop() || top10.getTop(),
        top10.isSetPercent() && top10.getPercent(),
        top10.getVal(),
        Optional.ofNullable(top10.isSetFilterVal() ? top10.getFilterVal() : null));
  }

  /** Reads the effective differential color used by a color filter. */
  public static AutofilterFilterCriterionReport colorCriterion(
      XSSFWorkbook workbook, CTColorFilter colorFilter) {
    return new AutofilterFilterCriterionReport.Color(
        colorFilter.isSetCellColor() && colorFilter.getCellColor(),
        Optional.ofNullable(
            colorFilter.isSetDxfId()
                ? dxfColor(
                        workbook,
                        colorFilter.getDxfId(),
                        colorFilter.isSetCellColor() && colorFilter.getCellColor())
                    .orElse(null)
                : null));
  }

  /** Reads the icon-set and icon index used by an icon filter. */
  public static AutofilterFilterCriterionReport iconCriterion(CTIconFilter iconFilter) {
    return new AutofilterFilterCriterionReport.Icon(
        iconFilter.getIconSet() == null ? "UNKNOWN" : iconFilter.getIconSet().toString(),
        iconFilter.isSetIconId() ? Math.toIntExact(iconFilter.getIconId()) : 0);
  }

  /** Reads one native sort condition without changing its state. */
  public static AutofilterSortConditionReport sortConditionSnapshot(
      XSSFWorkbook workbook, CTSortCondition sortCondition) {
    String range = Objects.requireNonNullElse(sortCondition.getRef(), "");
    if (range.isBlank()) {
      throw new IllegalArgumentException("autofilter sort condition is missing ref");
    }
    boolean descending = sortCondition.isSetDescending() && sortCondition.getDescending();
    STSortBy.Enum sortBy = sortCondition.getSortBy();
    if (!sortCondition.isSetSortBy() || sortBy == STSortBy.VALUE) {
      return new AutofilterSortConditionReport.Value(
          new dev.erst.gridgrind.core.vocabulary.RangeLabel(range), descending);
    }
    if (sortBy == STSortBy.CELL_COLOR) {
      return new AutofilterSortConditionReport.CellColor(
          new dev.erst.gridgrind.core.vocabulary.RangeLabel(range),
          descending,
          dxfColor(workbook, sortCondition.getDxfId(), true)
              .orElseThrow(
                  () ->
                      new IllegalArgumentException(
                          "autofilter cell-color sort condition is missing dxf color")));
    }
    if (sortBy == STSortBy.FONT_COLOR) {
      return new AutofilterSortConditionReport.FontColor(
          new dev.erst.gridgrind.core.vocabulary.RangeLabel(range),
          descending,
          dxfColor(workbook, sortCondition.getDxfId(), false)
              .orElseThrow(
                  () ->
                      new IllegalArgumentException(
                          "autofilter font-color sort condition is missing dxf color")));
    }
    return new AutofilterSortConditionReport.Icon(
        new dev.erst.gridgrind.core.vocabulary.RangeLabel(range),
        descending,
        Math.toIntExact(sortCondition.getIconId()));
  }

  private static STSortMethod.Enum toOoxmlSortMethod(ExcelAutofilterSortMethod sortMethod) {
    if (sortMethod == ExcelAutofilterSortMethod.PINYIN) {
      return STSortMethod.PIN_YIN;
    }
    return STSortMethod.STROKE;
  }

  /** Writes one authored filter criterion into the selected filter column. */
  public static void applyCriterion(
      XSSFWorkbook workbook, CTFilterColumn filterColumn, AutofilterFilterCriterionSpec criterion) {
    switch (criterion) {
      case AutofilterFilterCriterionSpec.Values values ->
          applyValuesCriterion(filterColumn, values);
      case AutofilterFilterCriterionSpec.Custom custom ->
          applyCustomCriterion(filterColumn, custom);
      case AutofilterFilterCriterionSpec.Dynamic dynamic ->
          applyDynamicCriterion(filterColumn, dynamic);
      case AutofilterFilterCriterionSpec.Top10 top10 -> applyTop10Criterion(filterColumn, top10);
      case AutofilterFilterCriterionSpec.Color color ->
          applyColorCriterion(workbook, filterColumn, color);
      case AutofilterFilterCriterionSpec.Icon icon -> applyIconCriterion(filterColumn, icon);
    }
  }

  private static void applyValuesCriterion(
      CTFilterColumn filterColumn, AutofilterFilterCriterionSpec.Values values) {
    CTFilters filters = filterColumn.addNewFilters();
    for (String value : values.values()) {
      filters.addNewFilter().setVal(value);
    }
    if (values.includeBlank()) {
      filters.setBlank(true);
    }
  }

  private static void applyCustomCriterion(
      CTFilterColumn filterColumn, AutofilterFilterCriterionSpec.Custom custom) {
    CTCustomFilters customFilters = filterColumn.addNewCustomFilters();
    customFilters.setAnd(custom.and());
    for (AutofilterFilterCriterionSpec.CustomConditionSpec condition : custom.conditions()) {
      CTCustomFilter customFilter = customFilters.addNewCustomFilter();
      customFilter.setOperator(requiredCustomFilterOperator(condition.operator()));
      customFilter.setVal(condition.value());
    }
  }

  private static STFilterOperator.Enum requiredCustomFilterOperator(String operator) {
    STFilterOperator.Enum resolved = STFilterOperator.Enum.forString(operator);
    if (resolved == null) {
      throw new IllegalArgumentException("unsupported autofilter custom operator: " + operator);
    }
    return resolved;
  }

  private static void applyDynamicCriterion(
      CTFilterColumn filterColumn, AutofilterFilterCriterionSpec.Dynamic dynamic) {
    CTDynamicFilter dynamicFilter = filterColumn.addNewDynamicFilter();
    dynamicFilter.setType(requiredDynamicFilterType(dynamic.dynamicType()));
    dynamic.value().ifPresent(dynamicFilter::setVal);
    dynamic.maxValue().ifPresent(dynamicFilter::setMaxVal);
  }

  private static STDynamicFilterType.Enum requiredDynamicFilterType(String type) {
    STDynamicFilterType.Enum resolved = STDynamicFilterType.Enum.forString(type);
    if (resolved == null) {
      throw new IllegalArgumentException("unsupported autofilter dynamic type: " + type);
    }
    return resolved;
  }

  private static void applyTop10Criterion(
      CTFilterColumn filterColumn, AutofilterFilterCriterionSpec.Top10 top10) {
    CTTop10 top10Filter = filterColumn.addNewTop10();
    top10Filter.setVal(top10.value());
    top10Filter.setTop(top10.top());
    top10Filter.setPercent(top10.percent());
  }

  private static void applyColorCriterion(
      XSSFWorkbook workbook,
      CTFilterColumn filterColumn,
      AutofilterFilterCriterionSpec.Color color) {
    CTColorFilter colorFilter = filterColumn.addNewColorFilter();
    colorFilter.setCellColor(color.cellColor());
    colorFilter.setDxfId(putColorDxf(workbook, color.color(), color.cellColor()) - 1L);
  }

  private static void applyIconCriterion(
      CTFilterColumn filterColumn, AutofilterFilterCriterionSpec.Icon icon) {
    CTIconFilter iconFilter = filterColumn.addNewIconFilter();
    iconFilter.setIconSet(requiredIconSet(icon.iconSet()));
    iconFilter.setIconId(icon.iconId());
  }

  private static STIconSetType.Enum requiredIconSet(String iconSet) {
    STIconSetType.Enum resolved = STIconSetType.Enum.forString(iconSet);
    if (resolved == null) {
      throw new IllegalArgumentException("unsupported autofilter icon set: " + iconSet);
    }
    return resolved;
  }

  /** Registers a differential cell or font color and returns its style index. */
  public static long putColorDxf(XSSFWorkbook workbook, ColorSpec color, boolean cellColor) {
    CTDxf dxf = CTDxf.Factory.newInstance();
    if (cellColor) {
      CTPatternFill patternFill = dxf.addNewFill().addNewPatternFill();
      patternFill.setPatternType(STPatternType.SOLID);
      patternFill.addNewFgColor().set(ColorAdapter.toXssfColor(workbook, color).getCTColor());
    } else {
      CTFont font = dxf.addNewFont();
      font.addNewColor().set(ColorAdapter.toXssfColor(workbook, color).getCTColor());
    }
    return workbook.getStylesSource().putDxf(dxf);
  }

  /** Reads an optional cell or font color from a differential style. */
  public static Optional<CellColorReport> dxfColor(
      XSSFWorkbook workbook, long dxfId, boolean cellColor) {
    CTDxf dxf = dxfAt(workbook.getStylesSource(), dxfId).orElse(null);
    if (dxf == null) {
      return Optional.empty();
    }
    return preferredDxfColor(workbook, dxf, cellColor).or(() -> fallbackDxfColor(workbook, dxf));
  }

  private static Optional<CellColorReport> preferredDxfColor(
      XSSFWorkbook workbook, CTDxf dxf, boolean cellColor) {
    return cellColor ? fillColor(workbook, dxf) : fontColor(workbook, dxf);
  }

  private static Optional<CellColorReport> fallbackDxfColor(XSSFWorkbook workbook, CTDxf dxf) {
    return fillColor(workbook, dxf).or(() -> fontColor(workbook, dxf));
  }

  private static Optional<CellColorReport> fillColor(XSSFWorkbook workbook, CTDxf dxf) {
    if (!dxf.isSetFill()
        || !dxf.getFill().isSetPatternFill()
        || !dxf.getFill().getPatternFill().isSetFgColor()) {
      return Optional.empty();
    }
    return ColorAdapter.snapshot(workbook, dxf.getFill().getPatternFill().getFgColor());
  }

  private static Optional<CellColorReport> fontColor(XSSFWorkbook workbook, CTDxf dxf) {
    if (!dxf.isSetFont() || dxf.getFont().sizeOfColorArray() == 0) {
      return Optional.empty();
    }
    return ColorAdapter.snapshot(workbook, dxf.getFont().getColorArray(0));
  }

  /** Looks up a differential style by its native workbook index. */
  public static Optional<CTDxf> dxfAt(StylesTable stylesTable, long dxfId) {
    if (dxfId < 0L || dxfId >= stylesTable._getDXfsSize()) {
      return Optional.empty();
    }
    return Optional.of(stylesTable.getDxfAt(Math.toIntExact(dxfId)));
  }
}
