package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Replays sheet-owned autofilter state onto a copied sheet surface. */
final class ExcelSheetCopyAutofilterSupport {
  private ExcelSheetCopyAutofilterSupport() {}

  static void replaceAutofilter(
      Optional<AutofilterEntryReport.SheetOwned> sheetAutofilter, ExcelSheet targetSheet) {
    targetSheet.metadata().clearAutofilter();
    sheetAutofilter.ifPresent(value -> copyAutofilter(value, targetSheet));
  }

  static Optional<AutofilterEntryReport.SheetOwned> sheetOwnedAutofilter(
      List<AutofilterEntryReport> autofilters) {
    Objects.requireNonNull(autofilters, "autofilters must not be null");
    if (autofilters.isEmpty()) {
      return Optional.empty();
    }
    AutofilterEntryReport autofilter = autofilters.getFirst();
    return switch (autofilter) {
      case AutofilterEntryReport.SheetOwned sheetOwned -> Optional.of(sheetOwned);
      case AutofilterEntryReport.TableOwned _ ->
          throw new IllegalStateException(
              "sheetOwnedAutofilters must not return table-owned autofilter snapshots");
    };
  }

  static Optional<String> sheetOwnedAutofilterRange(List<AutofilterEntryReport> autofilters) {
    return sheetOwnedAutofilter(autofilters).map(value -> value.range().value());
  }

  static AutofilterSortConditionSpec copyableSortCondition(
      AutofilterSortConditionReport condition) {
    return switch (condition) {
      case AutofilterSortConditionReport.Value value ->
          new AutofilterSortConditionSpec.Value(
              new dev.erst.gridgrind.core.vocabulary.RangeRef(value.range().value()),
              value.descending());
      case AutofilterSortConditionReport.CellColor cellColor ->
          new AutofilterSortConditionSpec.CellColor(
              new dev.erst.gridgrind.core.vocabulary.RangeRef(cellColor.range().value()),
              cellColor.descending(),
              Objects.requireNonNull(
                  ExcelColorSupport.copyOf(cellColor.color()), "cell sort color must not be null"));
      case AutofilterSortConditionReport.FontColor fontColor ->
          new AutofilterSortConditionSpec.FontColor(
              new dev.erst.gridgrind.core.vocabulary.RangeRef(fontColor.range().value()),
              fontColor.descending(),
              Objects.requireNonNull(
                  ExcelColorSupport.copyOf(fontColor.color()), "font sort color must not be null"));
      case AutofilterSortConditionReport.Icon icon ->
          new AutofilterSortConditionSpec.Icon(
              new dev.erst.gridgrind.core.vocabulary.RangeRef(icon.range().value()),
              icon.descending(),
              icon.iconId());
    };
  }

  private static void copyAutofilter(
      AutofilterEntryReport.SheetOwned sheetAutofilter, ExcelSheet targetSheet) {
    targetSheet
        .metadata()
        .setAutofilter(
            sheetAutofilter.range().value(),
            sheetAutofilter.filterColumns().stream()
                .map(ExcelSheetCopyAutofilterSupport::copyableAutofilterColumn)
                .toList(),
            copyableSortState(sheetAutofilter.sortState()));
  }

  private static AutofilterFilterColumnSpec copyableAutofilterColumn(
      AutofilterFilterColumnReport filterColumn) {
    return new AutofilterFilterColumnSpec(
        filterColumn.columnId(),
        filterColumn.showButton(),
        copyableAutofilterCriterion(filterColumn.criterion()));
  }

  private static AutofilterFilterCriterionSpec copyableAutofilterCriterion(
      AutofilterFilterCriterionReport criterion) {
    return switch (criterion) {
      case AutofilterFilterCriterionReport.Values values ->
          new AutofilterFilterCriterionSpec.Values(values.values(), values.includeBlank());
      case AutofilterFilterCriterionReport.Custom custom ->
          new AutofilterFilterCriterionSpec.Custom(
              custom.and(),
              custom.conditions().stream()
                  .map(
                      condition ->
                          new AutofilterFilterCriterionSpec.CustomConditionSpec(
                              condition.operator(), condition.value()))
                  .toList());
      case AutofilterFilterCriterionReport.Dynamic dynamic ->
          new AutofilterFilterCriterionSpec.Dynamic(
              dynamic.dynamicType(),
              Optional.ofNullable(dynamic.value()),
              Optional.ofNullable(dynamic.maxValue()));
      case AutofilterFilterCriterionReport.Top10 top10 ->
          new AutofilterFilterCriterionSpec.Top10(
              (int) Math.round(top10.value()), top10.top(), top10.percent());
      case AutofilterFilterCriterionReport.Color color ->
          new AutofilterFilterCriterionSpec.Color(
              color.cellColor(),
              Objects.requireNonNull(
                  ExcelColorSupport.copyOf(color.color()), "autofilter color must not be null"));
      case AutofilterFilterCriterionReport.Icon icon ->
          new AutofilterFilterCriterionSpec.Icon(icon.iconSet(), icon.iconId());
    };
  }

  private static Optional<AutofilterSortStateSpec> copyableSortState(
      Optional<AutofilterSortStateReport> sortState) {
    return Objects.requireNonNull(sortState, "sortState must not be null")
        .map(
            snapshot ->
                new AutofilterSortStateSpec(
                    new dev.erst.gridgrind.core.vocabulary.RangeRef(snapshot.range().value()),
                    snapshot.caseSensitive(),
                    snapshot.columnSort(),
                    snapshot.sortMethod(),
                    snapshot.conditions().stream()
                        .map(ExcelSheetCopyAutofilterSupport::copyableSortCondition)
                        .toList()));
  }
}
