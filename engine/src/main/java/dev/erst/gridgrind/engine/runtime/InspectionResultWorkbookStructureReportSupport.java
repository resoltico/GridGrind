package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.customxml.CustomXmlDataBindingReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlExportReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
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
}
