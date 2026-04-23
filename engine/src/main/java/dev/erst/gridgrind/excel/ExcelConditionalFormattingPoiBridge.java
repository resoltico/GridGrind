package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType;
import java.util.Objects;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;

/** Maps conditional-formatting enums between GridGrind and Apache POI. */
final class ExcelConditionalFormattingPoiBridge {
  private ExcelConditionalFormattingPoiBridge() {}

  static ExcelConditionalFormattingIconSet fromPoi(IconMultiStateFormatting.IconSet iconSet) {
    Objects.requireNonNull(iconSet, "iconSet must not be null");
    return ExcelConditionalFormattingIconSet.valueOf(iconSet.name());
  }

  static IconMultiStateFormatting.IconSet toPoi(ExcelConditionalFormattingIconSet iconSet) {
    Objects.requireNonNull(iconSet, "iconSet must not be null");
    return IconMultiStateFormatting.IconSet.valueOf(iconSet.name());
  }

  static ExcelConditionalFormattingThresholdType fromPoi(
      ConditionalFormattingThreshold.RangeType type) {
    Objects.requireNonNull(type, "type must not be null");
    return ExcelConditionalFormattingThresholdType.valueOf(type.name());
  }

  static ConditionalFormattingThreshold.RangeType toPoi(
      ExcelConditionalFormattingThresholdType type) {
    Objects.requireNonNull(type, "type must not be null");
    return ConditionalFormattingThreshold.RangeType.valueOf(type.name());
  }
}
