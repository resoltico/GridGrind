package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.ArrayFormulaReport;
import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellGradientStopReport;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.CustomXmlDataBindingReport;
import dev.erst.gridgrind.contract.dto.CustomXmlExportReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingReport;
import dev.erst.gridgrind.contract.dto.DrawingMarkerReport;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.GridGrindWorkbookSurfaceReports;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.TableColumnReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import java.util.List;

/** Workbook fact/report plain type descriptors. */
final class GridGrindProtocolCatalogWorkbookReportPlainTypeDescriptors {
  private GridGrindProtocolCatalogWorkbookReportPlainTypeDescriptors() {}

  static final List<CatalogPlainTypeDescriptor> DESCRIPTORS =
      List.of(
          plainTypeDescriptor(
              "workbookProtectionReportType",
              WorkbookProtectionReport.class,
              "WorkbookProtectionReport",
              "Exact workbook-protection report covering structure, windows, revisions,"
                  + " and password-hash presence flags.",
              List.of()),
          plainTypeDescriptor(
              "sheetSummaryReportType",
              GridGrindWorkbookSurfaceReports.SheetSummaryReport.class,
              "SheetSummaryReport",
              "Exact sheet summary report including visibility, protection, and structural counts.",
              List.of()),
          plainTypeDescriptor(
              "cellStyleReportType",
              GridGrindWorkbookSurfaceReports.CellStyleReport.class,
              "CellStyleReport",
              "Exact effective cell-style report used by style assertions.",
              List.of()),
          plainTypeDescriptor(
              "cellAlignmentReportType",
              CellAlignmentReport.class,
              "CellAlignmentReport",
              "Exact cell-alignment report.",
              List.of()),
          plainTypeDescriptor(
              "cellFontReportType",
              CellFontReport.class,
              "CellFontReport",
              "Exact cell-font report.",
              List.of("fontColor")),
          plainTypeDescriptor(
              "cellBorderReportType",
              CellBorderReport.class,
              "CellBorderReport",
              "Exact four-sided cell-border report.",
              List.of()),
          plainTypeDescriptor(
              "cellBorderSideReportType",
              CellBorderSideReport.class,
              "CellBorderSideReport",
              "Exact one-sided cell-border report.",
              List.of("color")),
          plainTypeDescriptor(
              "cellProtectionReportType",
              CellProtectionReport.class,
              "CellProtectionReport",
              "Exact cell-protection report.",
              List.of()),
          plainTypeDescriptor(
              "fontHeightReportType",
              FontHeightReport.class,
              "FontHeightReport",
              "Exact font-height report expressed in twips and points.",
              List.of()),
          plainTypeDescriptor(
              "cellGradientStopReportType",
              CellGradientStopReport.class,
              "CellGradientStopReport",
              "Exact gradient stop report.",
              List.of()),
          plainTypeDescriptor(
              "tableEntryReportType",
              TableEntryReport.class,
              "TableEntryReport",
              "Exact workbook table report used by table-facts assertions.",
              List.of("comment", "headerRowCellStyle", "dataCellStyle", "totalsRowCellStyle")),
          plainTypeDescriptor(
              "tableColumnReportType",
              TableColumnReport.class,
              "TableColumnReport",
              "Exact table-column report.",
              List.of(
                  "uniqueName", "totalsRowLabel", "totalsRowFunction", "calculatedColumnFormula")),
          plainTypeDescriptor(
              "drawingMarkerReportType",
              DrawingMarkerReport.class,
              "DrawingMarkerReport",
              "Exact cell-relative drawing marker report.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableAnchorReportType",
              PivotTableReport.Anchor.class,
              "PivotTableAnchorReport",
              "Exact pivot-table anchor report.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableFieldReportType",
              PivotTableReport.Field.class,
              "PivotTableFieldReport",
              "Exact pivot field report bound to one source column.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableDataFieldReportType",
              PivotTableReport.DataField.class,
              "PivotTableDataFieldReport",
              "Exact pivot data-field report.",
              List.of("valueFormat")),
          plainTypeDescriptor(
              "arrayFormulaReportType",
              ArrayFormulaReport.class,
              "ArrayFormulaReport",
              "One factual array-formula group report returned by GET_ARRAY_FORMULAS.",
              List.of()),
          plainTypeDescriptor(
              "customXmlMappingReportType",
              CustomXmlMappingReport.class,
              "CustomXmlMappingReport",
              "One factual workbook custom-XML mapping report.",
              List.of(
                  "schemaNamespace",
                  "schemaLanguage",
                  "schemaReference",
                  "schemaXml",
                  "dataBinding")),
          plainTypeDescriptor(
              "customXmlDataBindingReportType",
              CustomXmlDataBindingReport.class,
              "CustomXmlDataBindingReport",
              "Optional custom-XML data-binding metadata attached to one workbook mapping.",
              List.of("dataBindingName", "fileBinding", "connectionId", "fileBindingName")),
          plainTypeDescriptor(
              "customXmlLinkedCellReportType",
              CustomXmlLinkedCellReport.class,
              "CustomXmlLinkedCellReport",
              "One single-cell binding linked to a custom-XML mapping.",
              List.of()),
          plainTypeDescriptor(
              "customXmlLinkedTableReportType",
              CustomXmlLinkedTableReport.class,
              "CustomXmlLinkedTableReport",
              "One XML-mapped table linked to a custom-XML mapping.",
              List.of()),
          plainTypeDescriptor(
              "customXmlExportReportType",
              CustomXmlExportReport.class,
              "CustomXmlExportReport",
              "One exported custom-XML mapping payload plus the factual mapping metadata used to"
                  + " produce it.",
              List.of()),
          plainTypeDescriptor(
              "chartReportType",
              ChartReport.class,
              "ChartReport",
              "One factual chart report with chart-level presentation state and one or more"
                  + " plots.",
              List.of()),
          plainTypeDescriptor(
              "chartAxisReportType",
              ChartReport.Axis.class,
              "ChartAxisReport",
              "Exact chart-axis report.",
              List.of()),
          plainTypeDescriptor(
              "chartSeriesReportType",
              ChartReport.Series.class,
              "ChartSeriesReport",
              "Exact chart-series report."
                  + " smooth, marker, and explosion fields are populated only when the"
                  + " stored plot family supports them.",
              List.of("smooth", "markerStyle", "markerSize", "explosion")));

  private static CatalogPlainTypeDescriptor plainTypeDescriptor(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields) {
    return GridGrindProtocolCatalog.plainTypeDescriptor(
        group, recordType, id, summary, optionalFields);
  }
}
