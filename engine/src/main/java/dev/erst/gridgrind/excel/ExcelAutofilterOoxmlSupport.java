package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelAutofilterSortMethod;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.model.StylesTable;
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

/** Owns SpreadsheetML autofilter and sort-state XML translation. */
final class ExcelAutofilterOoxmlSupport {
  private ExcelAutofilterOoxmlSupport() {}

  static void replaceFilterColumns(
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

  static void replaceSortState(
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
    sortState
        .sortMethod()
        .ifPresent(sortMethod -> ctSortState.setSortMethod(toOoxmlSortMethod(sortMethod)));
  }

  private static void applySortCondition(
      XSSFWorkbook workbook, CTSortState ctSortState, ExcelAutofilterSortCondition condition) {
    CTSortCondition sortCondition = ctSortState.addNewSortCondition();
    sortCondition.setRef(condition.range());
    if (condition.descending()) {
      sortCondition.setDescending(true);
    }
    switch (condition) {
      case ExcelAutofilterSortCondition.Value _ -> {
        // SpreadsheetML uses VALUE semantics when no explicit sortBy discriminator is present.
      }
      case ExcelAutofilterSortCondition.CellColor cellColor -> {
        sortCondition.setSortBy(STSortBy.CELL_COLOR);
        sortCondition.setDxfId(putColorDxf(workbook, cellColor.color(), true) - 1L);
      }
      case ExcelAutofilterSortCondition.FontColor fontColor -> {
        sortCondition.setSortBy(STSortBy.FONT_COLOR);
        sortCondition.setDxfId(putColorDxf(workbook, fontColor.color(), false) - 1L);
      }
      case ExcelAutofilterSortCondition.Icon icon -> {
        sortCondition.setSortBy(STSortBy.ICON);
        sortCondition.setIconId(icon.iconId());
      }
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

  static List<ExcelAutofilterFilterColumnSnapshot> filterColumns(
      XSSFWorkbook workbook, CTAutoFilter autoFilter) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(autoFilter, "autoFilter must not be null");
    return java.util.Arrays.stream(autoFilter.getFilterColumnArray())
        .map(filterColumn -> filterColumnSnapshot(workbook, filterColumn))
        .toList();
  }

  static Optional<ExcelAutofilterSortStateSnapshot> sortState(
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
        new ExcelAutofilterSortStateSnapshot(
            range,
            sortState.isSetCaseSensitive() && sortState.getCaseSensitive(),
            sortState.isSetColumnSort() && sortState.getColumnSort(),
            sortState.isSetSortMethod()
                ? ExcelAutofilterSortMethod.fromOoxmlValue(sortState.getSortMethod().toString())
                : Optional.empty(),
            java.util.Arrays.stream(sortState.getSortConditionArray())
                .map(condition -> sortConditionSnapshot(workbook, condition))
                .toList()));
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
        java.util.Arrays.stream(customFilters.getCustomFilterArray())
            .map(ExcelAutofilterOoxmlSupport::customCondition)
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
                .orElse(null)
            : null);
  }

  static ExcelAutofilterFilterCriterionSnapshot iconCriterion(CTIconFilter iconFilter) {
    return new ExcelAutofilterFilterCriterionSnapshot.Icon(
        iconFilter.getIconSet() == null ? "UNKNOWN" : iconFilter.getIconSet().toString(),
        iconFilter.isSetIconId() ? Math.toIntExact(iconFilter.getIconId()) : 0);
  }

  static ExcelAutofilterSortConditionSnapshot sortConditionSnapshot(
      XSSFWorkbook workbook, CTSortCondition sortCondition) {
    String range = Objects.requireNonNullElse(sortCondition.getRef(), "");
    if (range.isBlank()) {
      throw new IllegalArgumentException("autofilter sort condition is missing ref");
    }
    boolean descending = sortCondition.isSetDescending() && sortCondition.getDescending();
    if (!sortCondition.isSetSortBy() || sortCondition.getSortBy() == STSortBy.VALUE) {
      return new ExcelAutofilterSortConditionSnapshot.Value(range, descending);
    }
    if (sortCondition.getSortBy() == STSortBy.CELL_COLOR) {
      return new ExcelAutofilterSortConditionSnapshot.CellColor(
          range,
          descending,
          dxfColor(workbook, sortCondition.getDxfId(), true)
              .orElseThrow(
                  () ->
                      new IllegalArgumentException(
                          "autofilter cell-color sort condition is missing dxf color")));
    }
    if (sortCondition.getSortBy() == STSortBy.FONT_COLOR) {
      return new ExcelAutofilterSortConditionSnapshot.FontColor(
          range,
          descending,
          dxfColor(workbook, sortCondition.getDxfId(), false)
              .orElseThrow(
                  () ->
                      new IllegalArgumentException(
                          "autofilter font-color sort condition is missing dxf color")));
    }
    if (sortCondition.getSortBy() == STSortBy.ICON) {
      return new ExcelAutofilterSortConditionSnapshot.Icon(
          range, descending, Math.toIntExact(sortCondition.getIconId()));
    }
    throw new IllegalArgumentException(
        "unsupported autofilter sortBy value: " + sortCondition.getSortBy().toString());
  }

  private static STSortMethod.Enum toOoxmlSortMethod(ExcelAutofilterSortMethod sortMethod) {
    return switch (sortMethod) {
      case PINYIN -> STSortMethod.PIN_YIN;
      case STROKE -> STSortMethod.STROKE;
    };
  }

  static Optional<ExcelColorSnapshot> dxfColor(
      XSSFWorkbook workbook, long dxfId, boolean cellColor) {
    CTDxf dxf = dxfAt(workbook.getStylesSource(), dxfId).orElse(null);
    if (dxf == null) {
      return Optional.empty();
    }
    if (cellColor
        && dxf.isSetFill()
        && dxf.getFill().isSetPatternFill()
        && dxf.getFill().getPatternFill().isSetFgColor()) {
      return Optional.of(
          ExcelColorSnapshotSupport.snapshot(
              workbook, dxf.getFill().getPatternFill().getFgColor()));
    }
    if (!cellColor && dxf.isSetFont() && dxf.getFont().sizeOfColorArray() > 0) {
      return Optional.of(
          ExcelColorSnapshotSupport.snapshot(workbook, dxf.getFont().getColorArray(0)));
    }
    if (dxf.isSetFill()
        && dxf.getFill().isSetPatternFill()
        && dxf.getFill().getPatternFill().isSetFgColor()) {
      return Optional.of(
          ExcelColorSnapshotSupport.snapshot(
              workbook, dxf.getFill().getPatternFill().getFgColor()));
    }
    if (dxf.isSetFont() && dxf.getFont().sizeOfColorArray() > 0) {
      return Optional.of(
          ExcelColorSnapshotSupport.snapshot(workbook, dxf.getFont().getColorArray(0)));
    }
    return Optional.empty();
  }

  static Optional<CTDxf> dxfAt(StylesTable stylesTable, long dxfId) {
    if (dxfId < 0L || dxfId >= stylesTable._getDXfsSize()) {
      return Optional.empty();
    }
    return Optional.of(stylesTable.getDxfAt(Math.toIntExact(dxfId)));
  }
}
