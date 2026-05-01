package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.AutofilterEntryReport;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnReport;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionReport;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateReport;
import dev.erst.gridgrind.contract.dto.CustomXmlDataBindingReport;
import dev.erst.gridgrind.contract.dto.CustomXmlExportReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingReport;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.TableColumnReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.TableStyleReport;
import dev.erst.gridgrind.excel.ExcelArrayFormulaSnapshot;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterColumnSnapshot;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterCriterionSnapshot;
import dev.erst.gridgrind.excel.ExcelAutofilterSnapshot;
import dev.erst.gridgrind.excel.ExcelAutofilterSortConditionSnapshot;
import dev.erst.gridgrind.excel.ExcelAutofilterSortStateSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlDataBindingSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlExportSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlLinkedCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlLinkedTableSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlMappingSnapshot;
import dev.erst.gridgrind.excel.ExcelPivotTableSnapshot;
import dev.erst.gridgrind.excel.ExcelTableColumnSnapshot;
import dev.erst.gridgrind.excel.ExcelTableSnapshot;
import dev.erst.gridgrind.excel.ExcelTableStyleSnapshot;

/** Converts workbook structure snapshots such as XML mappings, filters, tables, and pivots. */
final class InspectionResultWorkbookStructureReportSupport {
  private InspectionResultWorkbookStructureReportSupport() {}

  static CustomXmlMappingReport toCustomXmlMappingReport(ExcelCustomXmlMappingSnapshot snapshot) {
    return new CustomXmlMappingReport(
        snapshot.mapId(),
        snapshot.name(),
        snapshot.rootElement(),
        snapshot.schemaId(),
        snapshot.showImportExportValidationErrors(),
        snapshot.autoFit(),
        snapshot.append(),
        snapshot.preserveSortAfLayout(),
        snapshot.preserveFormat(),
        snapshot.schemaNamespace(),
        snapshot.schemaLanguage(),
        snapshot.schemaReference(),
        snapshot.schemaXml(),
        toCustomXmlDataBindingReport(snapshot.dataBinding()),
        snapshot.linkedCells().stream()
            .map(InspectionResultWorkbookStructureReportSupport::toCustomXmlLinkedCellReport)
            .toList(),
        snapshot.linkedTables().stream()
            .map(InspectionResultWorkbookStructureReportSupport::toCustomXmlLinkedTableReport)
            .toList());
  }

  static CustomXmlDataBindingReport toCustomXmlDataBindingReport(
      ExcelCustomXmlDataBindingSnapshot snapshot) {
    return snapshot == null
        ? null
        : new CustomXmlDataBindingReport(
            snapshot.dataBindingName(),
            snapshot.fileBinding(),
            snapshot.connectionId(),
            snapshot.fileBindingName(),
            snapshot.loadMode());
  }

  static CustomXmlLinkedCellReport toCustomXmlLinkedCellReport(
      ExcelCustomXmlLinkedCellSnapshot snapshot) {
    return new CustomXmlLinkedCellReport(
        snapshot.sheetName(), snapshot.address(), snapshot.xpath(), snapshot.xmlDataType());
  }

  static CustomXmlLinkedTableReport toCustomXmlLinkedTableReport(
      ExcelCustomXmlLinkedTableSnapshot snapshot) {
    return new CustomXmlLinkedTableReport(
        snapshot.sheetName(),
        snapshot.tableName(),
        snapshot.tableDisplayName(),
        snapshot.range(),
        snapshot.commonXPath());
  }

  static CustomXmlExportReport toCustomXmlExportReport(ExcelCustomXmlExportSnapshot snapshot) {
    return new CustomXmlExportReport(
        toCustomXmlMappingReport(snapshot.mapping()),
        snapshot.encoding(),
        snapshot.schemaValidated(),
        snapshot.xml());
  }

  static dev.erst.gridgrind.contract.dto.ArrayFormulaReport toArrayFormulaReport(
      ExcelArrayFormulaSnapshot snapshot) {
    return new dev.erst.gridgrind.contract.dto.ArrayFormulaReport(
        snapshot.sheetName(),
        snapshot.range(),
        snapshot.topLeftAddress(),
        snapshot.formula(),
        snapshot.singleCell());
  }

  static AutofilterEntryReport toAutofilterEntryReport(ExcelAutofilterSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelAutofilterSnapshot.SheetOwned sheetOwned ->
          new AutofilterEntryReport.SheetOwned(
              sheetOwned.range(),
              sheetOwned.filterColumns().stream()
                  .map(
                      InspectionResultWorkbookStructureReportSupport
                          ::toAutofilterFilterColumnReport)
                  .toList(),
              sheetOwned
                  .sortState()
                  .map(InspectionResultWorkbookStructureReportSupport::toAutofilterSortStateReport)
                  .orElse(null));
      case ExcelAutofilterSnapshot.TableOwned tableOwned ->
          new AutofilterEntryReport.TableOwned(
              tableOwned.range(),
              tableOwned.tableName(),
              tableOwned.filterColumns().stream()
                  .map(
                      InspectionResultWorkbookStructureReportSupport
                          ::toAutofilterFilterColumnReport)
                  .toList(),
              tableOwned
                  .sortState()
                  .map(InspectionResultWorkbookStructureReportSupport::toAutofilterSortStateReport)
                  .orElse(null));
    };
  }

  static TableEntryReport toTableEntryReport(ExcelTableSnapshot snapshot) {
    return new TableEntryReport(
        snapshot.name(),
        snapshot.sheetName(),
        snapshot.range(),
        snapshot.headerRowCount(),
        snapshot.totalsRowCount(),
        snapshot.columnNames(),
        snapshot.columns().stream()
            .map(InspectionResultWorkbookStructureReportSupport::toTableColumnReport)
            .toList(),
        toTableStyleReport(snapshot.style()),
        snapshot.hasAutofilter(),
        optionalText(snapshot.comment()),
        snapshot.published(),
        snapshot.insertRow(),
        snapshot.insertRowShift(),
        optionalText(snapshot.headerRowCellStyle()),
        optionalText(snapshot.dataCellStyle()),
        optionalText(snapshot.totalsRowCellStyle()));
  }

  static PivotTableReport toPivotTableReport(ExcelPivotTableSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelPivotTableSnapshot.Supported supported ->
          new PivotTableReport.Supported(
              supported.name(),
              supported.sheetName(),
              new PivotTableReport.Anchor(
                  supported.anchor().topLeftAddress(), supported.anchor().locationRange()),
              toPivotTableSourceReport(supported.source()),
              supported.rowLabels().stream()
                  .map(
                      field ->
                          new PivotTableReport.Field(
                              field.sourceColumnIndex(), field.sourceColumnName()))
                  .toList(),
              supported.columnLabels().stream()
                  .map(
                      field ->
                          new PivotTableReport.Field(
                              field.sourceColumnIndex(), field.sourceColumnName()))
                  .toList(),
              supported.reportFilters().stream()
                  .map(
                      field ->
                          new PivotTableReport.Field(
                              field.sourceColumnIndex(), field.sourceColumnName()))
                  .toList(),
              supported.dataFields().stream()
                  .map(
                      dataField ->
                          new PivotTableReport.DataField(
                              dataField.sourceColumnIndex(),
                              dataField.sourceColumnName(),
                              dataField.function(),
                              dataField.displayName(),
                              dataField.valueFormat()))
                  .toList(),
              supported.valuesAxisOnColumns());
      case ExcelPivotTableSnapshot.Unsupported unsupported ->
          new PivotTableReport.Unsupported(
              unsupported.name(),
              unsupported.sheetName(),
              new PivotTableReport.Anchor(
                  unsupported.anchor().topLeftAddress(), unsupported.anchor().locationRange()),
              unsupported.detail());
    };
  }

  private static PivotTableReport.Source toPivotTableSourceReport(
      ExcelPivotTableSnapshot.Source source) {
    return switch (source) {
      case ExcelPivotTableSnapshot.Source.Range range ->
          new PivotTableReport.Source.Range(range.sheetName(), range.range());
      case ExcelPivotTableSnapshot.Source.NamedRange namedRange ->
          new PivotTableReport.Source.NamedRange(
              namedRange.name(), namedRange.sheetName(), namedRange.range());
      case ExcelPivotTableSnapshot.Source.Table table ->
          new PivotTableReport.Source.Table(table.name(), table.sheetName(), table.range());
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
          new AutofilterFilterCriterionReport.Color(
              color.cellColor(), toCellColorReport(color.color()));
      case ExcelAutofilterFilterCriterionSnapshot.Icon icon ->
          new AutofilterFilterCriterionReport.Icon(icon.iconSet(), icon.iconId());
    };
  }

  private static AutofilterSortStateReport toAutofilterSortStateReport(
      ExcelAutofilterSortStateSnapshot sortState) {
    return new AutofilterSortStateReport(
        sortState.range(),
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
          new AutofilterSortConditionReport.Value(value.range(), value.descending());
      case ExcelAutofilterSortConditionSnapshot.CellColor cellColor ->
          new AutofilterSortConditionReport.CellColor(
              cellColor.range(), cellColor.descending(), toCellColorReport(cellColor.color()));
      case ExcelAutofilterSortConditionSnapshot.FontColor fontColor ->
          new AutofilterSortConditionReport.FontColor(
              fontColor.range(), fontColor.descending(), toCellColorReport(fontColor.color()));
      case ExcelAutofilterSortConditionSnapshot.Icon icon ->
          new AutofilterSortConditionReport.Icon(icon.range(), icon.descending(), icon.iconId());
    };
  }

  private static TableColumnReport toTableColumnReport(ExcelTableColumnSnapshot column) {
    return new TableColumnReport(
        column.id(),
        column.name(),
        optionalText(column.uniqueName()),
        optionalText(column.totalsRowLabel()),
        optionalText(column.totalsRowFunction()),
        optionalText(column.calculatedColumnFormula()));
  }

  private static TableStyleReport toTableStyleReport(ExcelTableStyleSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelTableStyleSnapshot.None _ -> new TableStyleReport.None();
      case ExcelTableStyleSnapshot.Named named ->
          new TableStyleReport.Named(
              named.name(),
              named.showFirstColumn(),
              named.showLastColumn(),
              named.showRowStripes(),
              named.showColumnStripes());
    };
  }

  private static dev.erst.gridgrind.contract.dto.CellColorReport toCellColorReport(
      dev.erst.gridgrind.excel.ExcelColorSnapshot color) {
    return InspectionResultCellReportSupport.toCellColorReport(color).orElse(null);
  }

  private static java.util.Optional<String> optionalText(String value) {
    if (value.isBlank()) {
      return java.util.Optional.empty();
    }
    return java.util.Optional.of(value);
  }
}
