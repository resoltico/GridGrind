package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColorFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCustomFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCustomFilters;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDynamicFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilterColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilters;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTIconFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTop10;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDynamicFilterType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STFilterOperator;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STIconSetType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType;

/** Writes authored autofilter criteria into SpreadsheetML filter-column XML. */
final class ExcelAutofilterCriterionWriteSupport {
  private ExcelAutofilterCriterionWriteSupport() {}

  static void applyCriterion(
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

  static long putColorDxf(XSSFWorkbook workbook, ColorSpec color, boolean cellColor) {
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
}
