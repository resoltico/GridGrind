package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.customxml.CustomXmlDataBindingReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlExportReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterColumnSnapshot;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterCriterionSnapshot;
import dev.erst.gridgrind.excel.ExcelAutofilterSnapshot;
import dev.erst.gridgrind.excel.ExcelAutofilterSortConditionSnapshot;
import dev.erst.gridgrind.excel.ExcelAutofilterSortStateSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlDataBindingSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlExportSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlLinkedCellSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlLinkedTableSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlMappingSnapshot;

/** Converts workbook structure snapshots such as XML mappings, filters, tables, and pivots. */
final class InspectionResultWorkbookStructureReportSupport {
  private InspectionResultWorkbookStructureReportSupport() {}

  static CustomXmlMappingReport toCustomXmlMappingReport(ExcelCustomXmlMappingSnapshot snapshot) {
    return new CustomXmlMappingReport(
        snapshot.mapId(),
        snapshot.name(),
        snapshot.rootElement(),
        snapshot.schemaId(),
        new CustomXmlMappingReport.Settings(
            snapshot.settings().showImportExportValidationErrors(),
            snapshot.settings().autoFit(),
            snapshot.settings().append(),
            snapshot.settings().preserveSortAfLayout(),
            snapshot.settings().preserveFormat()),
        new CustomXmlMappingReport.Schema(
            snapshot.schema().namespace().orElse(null),
            snapshot.schema().language().orElse(null),
            snapshot.schema().reference().orElse(null),
            snapshot.schema().xml().orElse(null)),
        snapshot
            .dataBinding()
            .map(InspectionResultWorkbookStructureReportSupport::toCustomXmlDataBindingReport)
            .orElse(null),
        snapshot.linkedCells().stream()
            .map(InspectionResultWorkbookStructureReportSupport::toCustomXmlLinkedCellReport)
            .toList(),
        snapshot.linkedTables().stream()
            .map(InspectionResultWorkbookStructureReportSupport::toCustomXmlLinkedTableReport)
            .toList());
  }

  static CustomXmlDataBindingReport toCustomXmlDataBindingReport(
      ExcelCustomXmlDataBindingSnapshot snapshot) {
    return new CustomXmlDataBindingReport(
        snapshot.dataBindingName().orElse(null),
        snapshot.fileBinding().orElse(null),
        snapshot.connectionId().orElse(null),
        snapshot.fileBindingName().orElse(null),
        snapshot.loadMode());
  }

  static CustomXmlLinkedCellReport toCustomXmlLinkedCellReport(
      ExcelCustomXmlLinkedCellSnapshot snapshot) {
    return new CustomXmlLinkedCellReport(
        new SheetLabel(snapshot.sheetName()),
        new CellAddress(snapshot.address()),
        snapshot.xpath(),
        snapshot.xmlDataType());
  }

  static CustomXmlLinkedTableReport toCustomXmlLinkedTableReport(
      ExcelCustomXmlLinkedTableSnapshot snapshot) {
    return new CustomXmlLinkedTableReport(
        new SheetLabel(snapshot.sheetName()),
        snapshot.tableName(),
        snapshot.tableDisplayName(),
        new RangeLabel(snapshot.range()),
        snapshot.commonXPath());
  }

  static CustomXmlExportReport toCustomXmlExportReport(ExcelCustomXmlExportSnapshot snapshot) {
    return new CustomXmlExportReport(
        toCustomXmlMappingReport(snapshot.mapping()),
        snapshot.encoding(),
        snapshot.schemaValidated(),
        snapshot.xml());
  }

  static dev.erst.gridgrind.core.model.formula.ArrayFormulaReport toArrayFormulaReport(
      dev.erst.gridgrind.excel.ExcelArrayFormulaSnapshot snapshot) {
    return new dev.erst.gridgrind.core.model.formula.ArrayFormulaReport(
        new SheetLabel(snapshot.sheetName()),
        new RangeLabel(snapshot.range()),
        new CellAddress(snapshot.topLeftAddress()),
        snapshot.formula(),
        snapshot.singleCell());
  }

  static AutofilterEntryReport toAutofilterEntryReport(ExcelAutofilterSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelAutofilterSnapshot.SheetOwned sheetOwned ->
          new AutofilterEntryReport.SheetOwned(
              new RangeLabel(sheetOwned.range()),
              sheetOwned.filterColumns().stream()
                  .map(
                      InspectionResultWorkbookStructureReportSupport
                          ::toAutofilterFilterColumnReport)
                  .toList(),
              sheetOwned
                  .sortState()
                  .map(
                      InspectionResultWorkbookStructureReportSupport::toAutofilterSortStateReport));
      case ExcelAutofilterSnapshot.TableOwned tableOwned ->
          new AutofilterEntryReport.TableOwned(
              new RangeLabel(tableOwned.range()),
              tableOwned.tableName(),
              tableOwned.filterColumns().stream()
                  .map(
                      InspectionResultWorkbookStructureReportSupport
                          ::toAutofilterFilterColumnReport)
                  .toList(),
              tableOwned
                  .sortState()
                  .map(
                      InspectionResultWorkbookStructureReportSupport::toAutofilterSortStateReport));
    };
  }

  private static AutofilterFilterColumnReport toAutofilterFilterColumnReport(
      ExcelAutofilterFilterColumnSnapshot filterColumn) {
    return new AutofilterFilterColumnReport(
        filterColumn.columnId(),
        filterColumn.showButton(),
        toAutofilterFilterCriterionReport(filterColumn.criterion()));
  }

  private static AutofilterFilterCriterionReport toAutofilterFilterCriterionReport(
      ExcelAutofilterFilterCriterionSnapshot criterion) {
    return switch (criterion) {
      case ExcelAutofilterFilterCriterionSnapshot.Values values ->
          new AutofilterFilterCriterionReport.Values(values.values(), values.includeBlank());
      case ExcelAutofilterFilterCriterionSnapshot.Custom custom ->
          new AutofilterFilterCriterionReport.Custom(
              custom.and(),
              custom.conditions().stream()
                  .map(
                      condition ->
                          new AutofilterFilterCriterionReport.CustomConditionReport(
                              condition.operator(), condition.value()))
                  .toList());
      case ExcelAutofilterFilterCriterionSnapshot.Dynamic dynamic ->
          new AutofilterFilterCriterionReport.Dynamic(
              dynamic.type(), dynamic.value(), dynamic.maxValue());
      case ExcelAutofilterFilterCriterionSnapshot.Top10 top10 ->
          new AutofilterFilterCriterionReport.Top10(
              top10.top(), top10.percent(), top10.value(), top10.filterValue());
      case ExcelAutofilterFilterCriterionSnapshot.Color color ->
          new AutofilterFilterCriterionReport.Color(color.cellColor(), color.color());
      case ExcelAutofilterFilterCriterionSnapshot.Icon icon ->
          new AutofilterFilterCriterionReport.Icon(icon.iconSet(), icon.iconId());
    };
  }

  private static AutofilterSortStateReport toAutofilterSortStateReport(
      ExcelAutofilterSortStateSnapshot sortState) {
    return new AutofilterSortStateReport(
        new RangeLabel(sortState.range()),
        sortState.caseSensitive(),
        sortState.columnSort(),
        sortState.sortMethod(),
        sortState.conditions().stream()
            .map(InspectionResultWorkbookStructureReportSupport::toAutofilterSortConditionReport)
            .toList());
  }

  private static AutofilterSortConditionReport toAutofilterSortConditionReport(
      ExcelAutofilterSortConditionSnapshot condition) {
    return switch (condition) {
      case ExcelAutofilterSortConditionSnapshot.Value value ->
          new AutofilterSortConditionReport.Value(
              new RangeLabel(value.range()), value.descending());
      case ExcelAutofilterSortConditionSnapshot.CellColor cellColor ->
          new AutofilterSortConditionReport.CellColor(
              new RangeLabel(cellColor.range()), cellColor.descending(), cellColor.color());
      case ExcelAutofilterSortConditionSnapshot.FontColor fontColor ->
          new AutofilterSortConditionReport.FontColor(
              new RangeLabel(fontColor.range()), fontColor.descending(), fontColor.color());
      case ExcelAutofilterSortConditionSnapshot.Icon icon ->
          new AutofilterSortConditionReport.Icon(
              new RangeLabel(icon.range()), icon.descending(), icon.iconId());
    };
  }
}
