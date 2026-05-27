package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Replays sheet-owned autofilter state onto a copied sheet surface. */
final class ExcelSheetCopyAutofilterSupport {
  private ExcelSheetCopyAutofilterSupport() {}

  static void replaceAutofilter(
      Optional<ExcelAutofilterSnapshot.SheetOwned> sheetAutofilter, ExcelSheet targetSheet) {
    targetSheet.metadata().clearAutofilter();
    sheetAutofilter.ifPresent(value -> copyAutofilter(value, targetSheet));
  }

  static Optional<ExcelAutofilterSnapshot.SheetOwned> sheetOwnedAutofilter(
      List<ExcelAutofilterSnapshot> autofilters) {
    Objects.requireNonNull(autofilters, "autofilters must not be null");
    if (autofilters.isEmpty()) {
      return Optional.empty();
    }
    ExcelAutofilterSnapshot autofilter = autofilters.getFirst();
    return switch (autofilter) {
      case ExcelAutofilterSnapshot.SheetOwned sheetOwned -> Optional.of(sheetOwned);
      case ExcelAutofilterSnapshot.TableOwned _ ->
          throw new IllegalStateException(
              "sheetOwnedAutofilters must not return table-owned autofilter snapshots");
    };
  }

  static Optional<String> sheetOwnedAutofilterRange(List<ExcelAutofilterSnapshot> autofilters) {
    return sheetOwnedAutofilter(autofilters).map(ExcelAutofilterSnapshot.SheetOwned::range);
  }

  static ExcelAutofilterSortCondition copyableSortCondition(
      ExcelAutofilterSortConditionSnapshot condition) {
    return switch (condition) {
      case ExcelAutofilterSortConditionSnapshot.Value value ->
          new ExcelAutofilterSortCondition.Value(value.range(), value.descending());
      case ExcelAutofilterSortConditionSnapshot.CellColor cellColor ->
          new ExcelAutofilterSortCondition.CellColor(
              cellColor.range(),
              cellColor.descending(),
              Objects.requireNonNull(
                  ExcelColorSupport.copyOf(cellColor.color()), "cell sort color must not be null"));
      case ExcelAutofilterSortConditionSnapshot.FontColor fontColor ->
          new ExcelAutofilterSortCondition.FontColor(
              fontColor.range(),
              fontColor.descending(),
              Objects.requireNonNull(
                  ExcelColorSupport.copyOf(fontColor.color()), "font sort color must not be null"));
      case ExcelAutofilterSortConditionSnapshot.Icon icon ->
          new ExcelAutofilterSortCondition.Icon(icon.range(), icon.descending(), icon.iconId());
    };
  }

  private static void copyAutofilter(
      ExcelAutofilterSnapshot.SheetOwned sheetAutofilter, ExcelSheet targetSheet) {
    targetSheet
        .metadata()
        .setAutofilter(
            sheetAutofilter.range(),
            sheetAutofilter.filterColumns().stream()
                .map(ExcelSheetCopyAutofilterSupport::copyableAutofilterColumn)
                .toList(),
            copyableSortState(sheetAutofilter.sortState()));
  }

  private static ExcelAutofilterFilterColumn copyableAutofilterColumn(
      ExcelAutofilterFilterColumnSnapshot filterColumn) {
    return new ExcelAutofilterFilterColumn(
        filterColumn.columnId(),
        filterColumn.showButton(),
        copyableAutofilterCriterion(filterColumn.criterion()));
  }

  private static ExcelAutofilterFilterCriterion copyableAutofilterCriterion(
      ExcelAutofilterFilterCriterionSnapshot criterion) {
    return switch (criterion) {
      case ExcelAutofilterFilterCriterionSnapshot.Values values ->
          new ExcelAutofilterFilterCriterion.Values(values.values(), values.includeBlank());
      case ExcelAutofilterFilterCriterionSnapshot.Custom custom ->
          new ExcelAutofilterFilterCriterion.Custom(
              custom.and(),
              custom.conditions().stream()
                  .map(
                      condition ->
                          new ExcelAutofilterFilterCriterion.CustomCondition(
                              condition.operator(), condition.value()))
                  .toList());
      case ExcelAutofilterFilterCriterionSnapshot.Dynamic dynamic ->
          new ExcelAutofilterFilterCriterion.Dynamic(
              dynamic.type(), dynamic.value(), dynamic.maxValue());
      case ExcelAutofilterFilterCriterionSnapshot.Top10 top10 ->
          new ExcelAutofilterFilterCriterion.Top10(
              (int) Math.round(top10.value()), top10.top(), top10.percent());
      case ExcelAutofilterFilterCriterionSnapshot.Color color ->
          new ExcelAutofilterFilterCriterion.Color(
              color.cellColor(),
              Objects.requireNonNull(
                  ExcelColorSupport.copyOf(color.color()), "autofilter color must not be null"));
      case ExcelAutofilterFilterCriterionSnapshot.Icon icon ->
          new ExcelAutofilterFilterCriterion.Icon(icon.iconSet(), icon.iconId());
    };
  }

  private static Optional<ExcelAutofilterSortState> copyableSortState(
      Optional<ExcelAutofilterSortStateSnapshot> sortState) {
    return Objects.requireNonNull(sortState, "sortState must not be null")
        .map(
            snapshot ->
                new ExcelAutofilterSortState(
                    snapshot.range(),
                    snapshot.caseSensitive(),
                    snapshot.columnSort(),
                    snapshot.sortMethod(),
                    snapshot.conditions().stream()
                        .map(ExcelSheetCopyAutofilterSupport::copyableSortCondition)
                        .toList()));
  }
}
