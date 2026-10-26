package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelAutofilterSortMethod;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColorFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCustomFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCustomFilters;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDynamicFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilterColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilters;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTIconFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSortCondition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSortState;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTop10;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSortBy;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSortMethod;

/** Owns SpreadsheetML autofilter and sort-state XML translation. */
final class ExcelAutofilterOoxmlSupport {
  private ExcelAutofilterOoxmlSupport() {}

  static void replaceFilterColumns(
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

  private static void applyCriterion(
      XSSFWorkbook workbook, CTFilterColumn filterColumn, AutofilterFilterCriterionSpec criterion) {
    ExcelAutofilterCriterionWriteSupport.applyCriterion(workbook, filterColumn, criterion);
  }

  static void replaceSortState(
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
        sortCondition.setDxfId(
            ExcelAutofilterCriterionWriteSupport.putColorDxf(workbook, cellColor.color(), true)
                - 1L);
      }
      case AutofilterSortConditionSpec.FontColor fontColor -> {
        sortCondition.setSortBy(STSortBy.FONT_COLOR);
        sortCondition.setDxfId(
            ExcelAutofilterCriterionWriteSupport.putColorDxf(workbook, fontColor.color(), false)
                - 1L);
      }
      case AutofilterSortConditionSpec.Icon icon -> {
        sortCondition.setSortBy(STSortBy.ICON);
        sortCondition.setIconId(icon.iconId());
      }
    }
  }

  static List<AutofilterFilterColumnReport> filterColumns(
      XSSFWorkbook workbook, CTAutoFilter autoFilter) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(autoFilter, "autoFilter must not be null");
    return java.util.Arrays.stream(autoFilter.getFilterColumnArray())
        .map(filterColumn -> filterColumnSnapshot(workbook, filterColumn))
        .toList();
  }

  static Optional<AutofilterSortStateReport> sortState(
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

  static AutofilterFilterCriterionReport valuesCriterion(CTFilters filters) {
    List<String> values = new ArrayList<>();
    for (CTFilter filter : filters.getFilterArray()) {
      values.add(Objects.requireNonNullElse(filter.getVal(), ""));
    }
    return new AutofilterFilterCriterionReport.Values(
        List.copyOf(values), filters.isSetBlank() && filters.getBlank());
  }

  static AutofilterFilterCriterionReport customCriterion(CTCustomFilters customFilters) {
    return new AutofilterFilterCriterionReport.Custom(
        customFilters.isSetAnd() && customFilters.getAnd(),
        java.util.Arrays.stream(customFilters.getCustomFilterArray())
            .map(ExcelAutofilterOoxmlSupport::customCondition)
            .toList());
  }

  private static AutofilterFilterCriterionReport.CustomConditionReport customCondition(
      CTCustomFilter customFilter) {
    String operator =
        customFilter.isSetOperator() ? customFilter.getOperator().toString() : "equal";
    return new AutofilterFilterCriterionReport.CustomConditionReport(
        operator, Objects.requireNonNullElse(customFilter.getVal(), ""));
  }

  static AutofilterFilterCriterionReport dynamicCriterion(CTDynamicFilter dynamicFilter) {
    return new AutofilterFilterCriterionReport.Dynamic(
        dynamicFilter.getType() == null ? "UNKNOWN" : dynamicFilter.getType().toString(),
        dynamicFilter.isSetVal() ? dynamicFilter.getVal() : null,
        dynamicFilter.isSetMaxVal() ? dynamicFilter.getMaxVal() : null);
  }

  static AutofilterFilterCriterionReport top10Criterion(CTTop10 top10) {
    return new AutofilterFilterCriterionReport.Top10(
        !top10.isSetTop() || top10.getTop(),
        top10.isSetPercent() && top10.getPercent(),
        top10.getVal(),
        top10.isSetFilterVal() ? top10.getFilterVal() : null);
  }

  static AutofilterFilterCriterionReport colorCriterion(
      XSSFWorkbook workbook, CTColorFilter colorFilter) {
    return new AutofilterFilterCriterionReport.Color(
        colorFilter.isSetCellColor() && colorFilter.getCellColor(),
        colorFilter.isSetDxfId()
            ? dxfColor(
                    workbook,
                    colorFilter.getDxfId(),
                    colorFilter.isSetCellColor() && colorFilter.getCellColor())
                .orElse(null)
            : null);
  }

  static AutofilterFilterCriterionReport iconCriterion(CTIconFilter iconFilter) {
    return new AutofilterFilterCriterionReport.Icon(
        iconFilter.getIconSet() == null ? "UNKNOWN" : iconFilter.getIconSet().toString(),
        iconFilter.isSetIconId() ? Math.toIntExact(iconFilter.getIconId()) : 0);
  }

  static AutofilterSortConditionReport sortConditionSnapshot(
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

  static Optional<CellColorReport> dxfColor(XSSFWorkbook workbook, long dxfId, boolean cellColor) {
    return ExcelAutofilterDxfColorSupport.dxfColor(workbook, dxfId, cellColor);
  }
}
