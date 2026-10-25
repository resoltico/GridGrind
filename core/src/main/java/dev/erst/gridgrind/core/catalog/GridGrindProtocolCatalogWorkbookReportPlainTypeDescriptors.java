package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlDataBindingReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlExportReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.model.drawing.CommentAnchorReport;
import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.layout.WindowDimensionsReport;
import dev.erst.gridgrind.core.model.layout.WindowRowReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellGradientStopReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.table.TableColumnReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.value.CellTemporalReport;
import dev.erst.gridgrind.core.result.SheetSummaryReport;
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
                  + " and password-hash presence flags."),
          plainTypeDescriptor(
              "sheetSummaryReportType",
              SheetSummaryReport.class,
              "SheetSummaryReport",
              "Exact sheet summary report including visibility, protection, and structural counts."),
          plainTypeDescriptor(
              "cellStyleReportType",
              CellStyleReport.class,
              "CellStyleReport",
              "Exact effective cell-style report used by style assertions."),
          plainTypeDescriptor(
              "cellAlignmentReportType",
              CellAlignmentReport.class,
              "CellAlignmentReport",
              "Exact cell-alignment report."),
          plainTypeDescriptor(
              "cellFontReportType",
              CellFontReport.class,
              "CellFontReport",
              "Exact cell-font report."),
          plainTypeDescriptor(
              "cellBorderReportType",
              CellBorderReport.class,
              "CellBorderReport",
              "Exact four-sided cell-border report."),
          plainTypeDescriptor(
              "cellProtectionReportType",
              CellProtectionReport.class,
              "CellProtectionReport",
              "Exact cell-protection report."),
          plainTypeDescriptor(
              "fontHeightReportType",
              FontHeightReport.class,
              "FontHeightReport",
              "Exact font-height report expressed in twips and points."),
          plainTypeDescriptor(
              "cellGradientStopReportType",
              CellGradientStopReport.class,
              "CellGradientStopReport",
              "Exact gradient stop report."),
          plainTypeDescriptor(
              "cellTemporalReportType",
              CellTemporalReport.class,
              "CellTemporalReport",
              "Projected temporal facet for one numeric cell value, including whether Excel"
                  + " formatting marks it as date-like and any derived DATE, TIME, or"
                  + " DATE_TIME semantic kind."),
          plainTypeDescriptor(
              "richTextRunReportType",
              RichTextRunReport.class,
              "RichTextRunReport",
              "One factual rich-text run reported from a text cell or comment."),
          plainTypeDescriptor(
              "commentAnchorReportType",
              CommentAnchorReport.class,
              "CommentAnchorReport",
              "Exact comment-anchor bounds reported from worksheet drawing metadata."),
          plainTypeDescriptor(
              "commentReportType",
              CommentReport.class,
              "CommentReport",
              "Exact factual comment report including optional rich-text runs and anchor bounds."),
          plainTypeDescriptor(
              "windowDimensionsReportType",
              WindowDimensionsReport.class,
              "WindowDimensionsReport",
              "Published row and column dimensions for one rectangular cell window."),
          plainTypeDescriptor(
              "windowRowReportType",
              WindowRowReport.class,
              "WindowRowReport",
              "One dense window row preserving row index and ordered cell payloads."),
          plainTypeDescriptor(
              "tableEntryReportType",
              TableEntryReport.class,
              "TableEntryReport",
              "Exact workbook table report used by table-facts assertions."),
          plainTypeDescriptor(
              "tableEntryStructureReportType",
              TableEntryReport.Structure.class,
              "TableEntryStructureReport",
              "Structural table facts including header/totals counts and per-column metadata."),
          plainTypeDescriptor(
              "tableEntryBehaviorReportType",
              TableEntryReport.Behavior.class,
              "TableEntryBehaviorReport",
              "Persisted workbook-table behavior toggles."),
          plainTypeDescriptor(
              "tableEntryPresentationReportType",
              TableEntryReport.Presentation.class,
              "TableEntryPresentationReport",
              "Optional table comment and style labels attached to one persisted workbook table."),
          plainTypeDescriptor(
              "tableColumnReportType",
              TableColumnReport.class,
              "TableColumnReport",
              "Exact table-column report."),
          plainTypeDescriptor(
              "drawingMarkerReportType",
              DrawingMarkerReport.class,
              "DrawingMarkerReport",
              "Exact cell-relative drawing marker report."),
          plainTypeDescriptor(
              "pivotTableAnchorReportType",
              PivotTableReport.Anchor.class,
              "PivotTableAnchorReport",
              "Exact pivot-table anchor report."),
          plainTypeDescriptor(
              "pivotTableFieldReportType",
              PivotTableReport.Field.class,
              "PivotTableFieldReport",
              "Exact pivot field report bound to one source column."),
          plainTypeDescriptor(
              "pivotTableDataFieldReportType",
              PivotTableReport.DataField.class,
              "PivotTableDataFieldReport",
              "Exact pivot data-field report."),
          plainTypeDescriptor(
              "arrayFormulaReportType",
              ArrayFormulaReport.class,
              "ArrayFormulaReport",
              "One factual array-formula group report returned by GET_ARRAY_FORMULAS."),
          plainTypeDescriptor(
              "customXmlMappingReportType",
              CustomXmlMappingReport.class,
              "CustomXmlMappingReport",
              "One factual workbook custom-XML mapping report."),
          plainTypeDescriptor(
              "customXmlMappingSettingsReportType",
              CustomXmlMappingReport.Settings.class,
              "CustomXmlMappingSettingsReport",
              "Persisted custom-XML map behavior flags."),
          plainTypeDescriptor(
              "customXmlMappingSchemaReportType",
              CustomXmlMappingReport.Schema.class,
              "CustomXmlMappingSchemaReport",
              "Optional schema metadata attached to one custom-XML mapping."),
          plainTypeDescriptor(
              "customXmlDataBindingReportType",
              CustomXmlDataBindingReport.class,
              "CustomXmlDataBindingReport",
              "Optional custom-XML data-binding metadata attached to one workbook mapping."),
          plainTypeDescriptor(
              "customXmlLinkedCellReportType",
              CustomXmlLinkedCellReport.class,
              "CustomXmlLinkedCellReport",
              "One single-cell binding linked to a custom-XML mapping."),
          plainTypeDescriptor(
              "customXmlLinkedTableReportType",
              CustomXmlLinkedTableReport.class,
              "CustomXmlLinkedTableReport",
              "One XML-mapped table linked to a custom-XML mapping."),
          plainTypeDescriptor(
              "customXmlExportReportType",
              CustomXmlExportReport.class,
              "CustomXmlExportReport",
              "One exported custom-XML mapping payload plus the factual mapping metadata used to"
                  + " produce it."),
          plainTypeDescriptor(
              "chartReportType",
              ChartReport.class,
              "ChartReport",
              "One factual chart report with chart-level presentation state and one or more"
                  + " plots."),
          plainTypeDescriptor(
              "chartAxisReportType",
              ChartReport.Axis.class,
              "ChartAxisReport",
              "Exact chart-axis report."),
          plainTypeDescriptor(
              "chartSeriesReportType",
              ChartReport.Series.class,
              "ChartSeriesReport",
              "Exact chart-series report."
                  + " smooth, marker, and explosion fields are populated only when the"
                  + " stored plot family supports them."));

  private static CatalogPlainTypeDescriptor plainTypeDescriptor(
      String group, Class<? extends Record> recordType, String id, String summary) {
    return CatalogTypeEntryFactory.plainTypeDescriptor(group, recordType, id, summary);
  }
}
