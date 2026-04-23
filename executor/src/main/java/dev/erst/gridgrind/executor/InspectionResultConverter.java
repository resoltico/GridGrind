package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.ArrayFormulaReport;
import dev.erst.gridgrind.contract.dto.AutofilterEntryReport;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnReport;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.contract.dto.AutofilterHealthReport;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionReport;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdReport;
import dev.erst.gridgrind.contract.dto.CustomXmlDataBindingReport;
import dev.erst.gridgrind.contract.dto.CustomXmlExportReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingReport;
import dev.erst.gridgrind.contract.dto.DataValidationEntryReport;
import dev.erst.gridgrind.contract.dto.DataValidationHealthReport;
import dev.erst.gridgrind.contract.dto.DifferentialBorderReport;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideReport;
import dev.erst.gridgrind.contract.dto.DifferentialStyleReport;
import dev.erst.gridgrind.contract.dto.DrawingAnchorReport;
import dev.erst.gridgrind.contract.dto.DrawingMarkerReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectPayloadReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectReport;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextReport;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.IgnoredErrorReport;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionReport;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureReport;
import dev.erst.gridgrind.contract.dto.PaneReport;
import dev.erst.gridgrind.contract.dto.PivotTableHealthReport;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.PrintAreaReport;
import dev.erst.gridgrind.contract.dto.PrintLayoutReport;
import dev.erst.gridgrind.contract.dto.PrintMarginsReport;
import dev.erst.gridgrind.contract.dto.PrintScalingReport;
import dev.erst.gridgrind.contract.dto.PrintSetupReport;
import dev.erst.gridgrind.contract.dto.PrintTitleColumnsReport;
import dev.erst.gridgrind.contract.dto.PrintTitleRowsReport;
import dev.erst.gridgrind.contract.dto.SheetDefaultsReport;
import dev.erst.gridgrind.contract.dto.SheetDisplayReport;
import dev.erst.gridgrind.contract.dto.SheetOutlineSummaryReport;
import dev.erst.gridgrind.contract.dto.SheetPresentationReport;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.TableColumnReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.TableHealthReport;
import dev.erst.gridgrind.contract.dto.TableStyleReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.excel.ExcelArrayFormulaSnapshot;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterColumnSnapshot;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterCriterionSnapshot;
import dev.erst.gridgrind.excel.ExcelAutofilterSnapshot;
import dev.erst.gridgrind.excel.ExcelAutofilterSortStateSnapshot;
import dev.erst.gridgrind.excel.ExcelBorderSideSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelChartSnapshot;
import dev.erst.gridgrind.excel.ExcelColorSnapshot;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelCommentSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlDataBindingSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlExportSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlLinkedCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlLinkedTableSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlMappingSnapshot;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationSnapshot;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialBorderSide;
import dev.erst.gridgrind.excel.ExcelDifferentialStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelDrawingObjectPayload;
import dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelHeaderFooterText;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelPivotTableSnapshot;
import dev.erst.gridgrind.excel.ExcelPrintLayout;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetPresentationSnapshot;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelTableColumnSnapshot;
import dev.erst.gridgrind.excel.ExcelTableSnapshot;
import dev.erst.gridgrind.excel.ExcelTableStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSnapshot;
import dev.erst.gridgrind.excel.WorkbookAnalysis;

/**
 * Converts workbook-core inspection results into protocol response shapes.
 *
 * <p>This translation seam intentionally spans the factual read surface on both sides.
 */
@SuppressWarnings("PMD.ExcessiveImports")
final class InspectionResultConverter {
  private InspectionResultConverter() {}

  /** Converts one workbook-core inspection result into the protocol inspection-result shape. */
  static InspectionResult toReadResult(dev.erst.gridgrind.excel.WorkbookReadResult result) {
    return switch (result) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult workbookSummary ->
          new InspectionResult.WorkbookSummaryResult(
              workbookSummary.stepId(), toWorkbookSummary(workbookSummary.workbook()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.PackageSecurityResult packageSecurity ->
          new InspectionResult.PackageSecurityResult(
              packageSecurity.stepId(), toOoxmlPackageSecurityReport(packageSecurity.security()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookProtectionResult protection ->
          new InspectionResult.WorkbookProtectionResult(
              protection.stepId(), toWorkbookProtectionReport(protection.protection()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.CustomXmlMappingsResult customXmlMappings ->
          new InspectionResult.CustomXmlMappingsResult(
              customXmlMappings.stepId(),
              customXmlMappings.mappings().stream()
                  .map(InspectionResultConverter::toCustomXmlMappingReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.CustomXmlExportResult customXmlExport ->
          new InspectionResult.CustomXmlExportResult(
              customXmlExport.stepId(), toCustomXmlExportReport(customXmlExport.export()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangesResult namedRanges ->
          new InspectionResult.NamedRangesResult(
              namedRanges.stepId(),
              namedRanges.namedRanges().stream()
                  .map(InspectionResultConverter::toNamedRangeReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult sheetSummary ->
          new InspectionResult.SheetSummaryResult(
              sheetSummary.stepId(),
              new GridGrindResponse.SheetSummaryReport(
                  sheetSummary.sheet().sheetName(),
                  sheetSummary.sheet().visibility(),
                  toSheetProtectionReport(sheetSummary.sheet().protection()),
                  sheetSummary.sheet().physicalRowCount(),
                  sheetSummary.sheet().lastRowIndex(),
                  sheetSummary.sheet().lastColumnIndex()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.ArrayFormulasResult arrayFormulas ->
          new InspectionResult.ArrayFormulasResult(
              arrayFormulas.stepId(),
              arrayFormulas.arrayFormulas().stream()
                  .map(InspectionResultConverter::toArrayFormulaReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.CellsResult cells ->
          new InspectionResult.CellsResult(
              cells.stepId(),
              cells.sheetName(),
              cells.cells().stream().map(InspectionResultConverter::toCellReport).toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.WindowResult window ->
          new InspectionResult.WindowResult(window.stepId(), toWindowReport(window.window()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.MergedRegionsResult mergedRegions ->
          new InspectionResult.MergedRegionsResult(
              mergedRegions.stepId(),
              mergedRegions.sheetName(),
              mergedRegions.mergedRegions().stream()
                  .map(region -> new GridGrindResponse.MergedRegionReport(region.range()))
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinksResult hyperlinks ->
          new InspectionResult.HyperlinksResult(
              hyperlinks.stepId(),
              hyperlinks.sheetName(),
              hyperlinks.hyperlinks().stream()
                  .map(InspectionResultConverter::toCellHyperlinkReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.CommentsResult comments ->
          new InspectionResult.CommentsResult(
              comments.stepId(),
              comments.sheetName(),
              comments.comments().stream()
                  .map(InspectionResultConverter::toCellCommentReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.DrawingObjectsResult drawingObjects ->
          new InspectionResult.DrawingObjectsResult(
              drawingObjects.stepId(),
              drawingObjects.sheetName(),
              drawingObjects.drawingObjects().stream()
                  .map(InspectionResultConverter::toDrawingObjectReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.ChartsResult charts ->
          new InspectionResult.ChartsResult(
              charts.stepId(),
              charts.sheetName(),
              charts.charts().stream().map(InspectionResultConverter::toChartReport).toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.PivotTablesResult pivotTables ->
          new InspectionResult.PivotTablesResult(
              pivotTables.stepId(),
              pivotTables.pivotTables().stream()
                  .map(InspectionResultConverter::toPivotTableReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.DrawingObjectPayloadResult drawingPayload ->
          new InspectionResult.DrawingObjectPayloadResult(
              drawingPayload.stepId(),
              drawingPayload.sheetName(),
              toDrawingObjectPayloadReport(drawingPayload.payload()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult sheetLayout ->
          new InspectionResult.SheetLayoutResult(
              sheetLayout.stepId(), toSheetLayoutReport(sheetLayout.layout()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult printLayout ->
          new InspectionResult.PrintLayoutResult(
              printLayout.stepId(), toPrintLayoutReport(printLayout));
      case dev.erst.gridgrind.excel.WorkbookReadResult.DataValidationsResult dataValidations ->
          new InspectionResult.DataValidationsResult(
              dataValidations.stepId(),
              dataValidations.sheetName(),
              dataValidations.validations().stream()
                  .map(InspectionResultConverter::toDataValidationEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult
              conditionalFormatting ->
          new InspectionResult.ConditionalFormattingResult(
              conditionalFormatting.stepId(),
              conditionalFormatting.sheetName(),
              conditionalFormatting.conditionalFormattingBlocks().stream()
                  .map(InspectionResultConverter::toConditionalFormattingEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.AutofiltersResult autofilters ->
          new InspectionResult.AutofiltersResult(
              autofilters.stepId(),
              autofilters.sheetName(),
              autofilters.autofilters().stream()
                  .map(InspectionResultConverter::toAutofilterEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.TablesResult tables ->
          new InspectionResult.TablesResult(
              tables.stepId(),
              tables.tables().stream().map(InspectionResultConverter::toTableEntryReport).toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurfaceResult formulaSurface ->
          new InspectionResult.FormulaSurfaceResult(
              formulaSurface.stepId(), toFormulaSurfaceReport(formulaSurface.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchemaResult sheetSchema ->
          new InspectionResult.SheetSchemaResult(
              sheetSchema.stepId(), toSheetSchemaReport(sheetSchema.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceResult namedRangeSurface ->
          new InspectionResult.NamedRangeSurfaceResult(
              namedRangeSurface.stepId(), toNamedRangeSurfaceReport(namedRangeSurface.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.FormulaHealthResult formulaHealth ->
          new InspectionResult.FormulaHealthResult(
              formulaHealth.stepId(), toFormulaHealthReport(formulaHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.DataValidationHealthResult
              dataValidationHealth ->
          new InspectionResult.DataValidationHealthResult(
              dataValidationHealth.stepId(),
              toDataValidationHealthReport(dataValidationHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingHealthResult
              conditionalFormattingHealth ->
          new InspectionResult.ConditionalFormattingHealthResult(
              conditionalFormattingHealth.stepId(),
              toConditionalFormattingHealthReport(conditionalFormattingHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.AutofilterHealthResult autofilterHealth ->
          new InspectionResult.AutofilterHealthResult(
              autofilterHealth.stepId(), toAutofilterHealthReport(autofilterHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.TableHealthResult tableHealth ->
          new InspectionResult.TableHealthResult(
              tableHealth.stepId(), toTableHealthReport(tableHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.PivotTableHealthResult pivotTableHealth ->
          new InspectionResult.PivotTableHealthResult(
              pivotTableHealth.stepId(), toPivotTableHealthReport(pivotTableHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinkHealthResult hyperlinkHealth ->
          new InspectionResult.HyperlinkHealthResult(
              hyperlinkHealth.stepId(), toHyperlinkHealthReport(hyperlinkHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeHealthResult namedRangeHealth ->
          new InspectionResult.NamedRangeHealthResult(
              namedRangeHealth.stepId(), toNamedRangeHealthReport(namedRangeHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookFindingsResult workbookFindings ->
          new InspectionResult.WorkbookFindingsResult(
              workbookFindings.stepId(), toWorkbookFindingsReport(workbookFindings.analysis()));
    };
  }

  /** Converts one workbook-core workbook summary into the protocol inspection-result shape. */
  static GridGrindResponse.WorkbookSummary toWorkbookSummary(
      dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary workbookSummary) {
    return switch (workbookSummary) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.Empty empty ->
          new GridGrindResponse.WorkbookSummary.Empty(
              empty.sheetCount(),
              empty.sheetNames(),
              empty.namedRangeCount(),
              empty.forceFormulaRecalculationOnOpen());
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.WithSheets withSheets ->
          new GridGrindResponse.WorkbookSummary.WithSheets(
              withSheets.sheetCount(),
              withSheets.sheetNames(),
              withSheets.activeSheetName(),
              withSheets.selectedSheetNames(),
              withSheets.namedRangeCount(),
              withSheets.forceFormulaRecalculationOnOpen());
    };
  }

  static OoxmlPackageSecurityReport toOoxmlPackageSecurityReport(
      dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySnapshot snapshot) {
    return new OoxmlPackageSecurityReport(
        toOoxmlEncryptionReport(snapshot.encryption()),
        snapshot.signatures().stream()
            .map(InspectionResultConverter::toOoxmlSignatureReport)
            .toList());
  }

  static OoxmlEncryptionReport toOoxmlEncryptionReport(
      dev.erst.gridgrind.excel.ExcelOoxmlEncryptionSnapshot snapshot) {
    return new OoxmlEncryptionReport(
        snapshot.encrypted(),
        snapshot.mode(),
        snapshot.cipherAlgorithm(),
        snapshot.hashAlgorithm(),
        snapshot.chainingMode(),
        snapshot.keyBits(),
        snapshot.blockSize(),
        snapshot.spinCount());
  }

  static OoxmlSignatureReport toOoxmlSignatureReport(
      dev.erst.gridgrind.excel.ExcelOoxmlSignatureSnapshot snapshot) {
    return new OoxmlSignatureReport(
        snapshot.packagePartName(),
        snapshot.signerSubject(),
        snapshot.signerIssuer(),
        snapshot.serialNumberHex(),
        snapshot.state());
  }

  /** Converts one workbook-core named-range snapshot into the protocol inspection-result shape. */
  static GridGrindResponse.NamedRangeReport toNamedRangeReport(ExcelNamedRangeSnapshot namedRange) {
    return switch (namedRange) {
      case ExcelNamedRangeSnapshot.RangeSnapshot rangeSnapshot ->
          new GridGrindResponse.NamedRangeReport.RangeReport(
              rangeSnapshot.name(),
              toNamedRangeScope(rangeSnapshot.scope()),
              rangeSnapshot.refersToFormula(),
              new NamedRangeTarget(
                  rangeSnapshot.target().sheetName(), rangeSnapshot.target().range()));
      case ExcelNamedRangeSnapshot.FormulaSnapshot formulaSnapshot ->
          new GridGrindResponse.NamedRangeReport.FormulaReport(
              formulaSnapshot.name(),
              toNamedRangeScope(formulaSnapshot.scope()),
              formulaSnapshot.refersToFormula());
    };
  }

  /** Converts the workbook-core named-range scope into the protocol scope variant. */
  static NamedRangeScope toNamedRangeScope(ExcelNamedRangeScope scope) {
    return switch (scope) {
      case ExcelNamedRangeScope.WorkbookScope _ -> new NamedRangeScope.Workbook();
      case ExcelNamedRangeScope.SheetScope sheetScope ->
          new NamedRangeScope.Sheet(sheetScope.sheetName());
    };
  }

  /** Converts workbook-core sheet-protection state into the protocol response variant. */
  static GridGrindResponse.SheetProtectionReport toSheetProtectionReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection protection) {
    return switch (protection) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection.Unprotected _ ->
          new GridGrindResponse.SheetProtectionReport.Unprotected();
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection.Protected protectedState ->
          new GridGrindResponse.SheetProtectionReport.Protected(
              toSheetProtectionSettings(protectedState.settings()));
    };
  }

  /** Converts workbook-core hyperlink metadata into the canonical protocol hyperlink shape. */
  static HyperlinkTarget toHyperlinkTarget(ExcelHyperlink hyperlink) {
    return InspectionResultCellReportSupport.toHyperlinkTarget(hyperlink);
  }

  /** Converts workbook-core comment metadata into the protocol inspection-result shape. */
  static GridGrindResponse.CommentReport toCommentReport(ExcelComment comment) {
    return InspectionResultCellReportSupport.toCommentReport(comment);
  }

  static DrawingObjectReport toDrawingObjectReport(ExcelDrawingObjectSnapshot snapshot) {
    return InspectionResultDrawingReportSupport.toDrawingObjectReport(snapshot);
  }

  static ChartReport toChartReport(ExcelChartSnapshot snapshot) {
    return InspectionResultDrawingReportSupport.toChartReport(snapshot);
  }

  static DrawingObjectPayloadReport toDrawingObjectPayloadReport(
      ExcelDrawingObjectPayload payload) {
    return InspectionResultDrawingReportSupport.toDrawingObjectPayloadReport(payload);
  }

  static DrawingAnchorReport toDrawingAnchorReport(ExcelDrawingAnchor anchor) {
    return InspectionResultDrawingReportSupport.toDrawingAnchorReport(anchor);
  }

  static DrawingMarkerReport toDrawingMarkerReport(ExcelDrawingMarker marker) {
    return InspectionResultDrawingReportSupport.toDrawingMarkerReport(marker);
  }

  /** Converts workbook-core workbook-protection state into the protocol inspection-result shape. */
  static WorkbookProtectionReport toWorkbookProtectionReport(
      ExcelWorkbookProtectionSnapshot protection) {
    return new WorkbookProtectionReport(
        protection.structureLocked(),
        protection.windowsLocked(),
        protection.revisionsLocked(),
        protection.workbookPasswordHashPresent(),
        protection.revisionsPasswordHashPresent());
  }

  static DataValidationEntryReport toDataValidationEntryReport(
      ExcelDataValidationSnapshot snapshot) {
    return InspectionResultValidationReportSupport.toDataValidationEntryReport(snapshot);
  }

  static DataValidationEntryReport.DataValidationDefinitionReport toDataValidationDefinitionReport(
      ExcelDataValidationDefinition definition) {
    return InspectionResultValidationReportSupport.toDataValidationDefinitionReport(definition);
  }

  static ConditionalFormattingEntryReport toConditionalFormattingEntryReport(
      ExcelConditionalFormattingBlockSnapshot block) {
    return InspectionResultValidationReportSupport.toConditionalFormattingEntryReport(block);
  }

  static ConditionalFormattingRuleReport toConditionalFormattingRuleReport(
      ExcelConditionalFormattingRuleSnapshot rule) {
    return InspectionResultValidationReportSupport.toConditionalFormattingRuleReport(rule);
  }

  static ConditionalFormattingThresholdReport toConditionalFormattingThresholdReport(
      ExcelConditionalFormattingThresholdSnapshot threshold) {
    return InspectionResultValidationReportSupport.toConditionalFormattingThresholdReport(
        threshold);
  }

  static DifferentialStyleReport toDifferentialStyleReport(ExcelDifferentialStyleSnapshot style) {
    return InspectionResultValidationReportSupport.toDifferentialStyleReport(style);
  }

  static DifferentialBorderReport toDifferentialBorderReport(ExcelDifferentialBorder border) {
    return InspectionResultValidationReportSupport.toDifferentialBorderReport(border);
  }

  static DifferentialBorderSideReport toDifferentialBorderSideReport(
      ExcelDifferentialBorderSide side) {
    return InspectionResultValidationReportSupport.toDifferentialBorderSideReport(side);
  }

  static FontHeightReport toFontHeightReport(ExcelFontHeight fontHeight) {
    return InspectionResultCellReportSupport.toFontHeightReport(fontHeight);
  }

  static GridGrindResponse.CellStyleReport toCellStyleReport(ExcelCellStyleSnapshot style) {
    return InspectionResultCellReportSupport.toCellStyleReport(style);
  }

  static CellFontReport toCellFontReport(ExcelCellFontSnapshot font) {
    return InspectionResultCellReportSupport.toCellFontReport(font);
  }

  static CellBorderSideReport toCellBorderSideReport(ExcelBorderSideSnapshot side) {
    return InspectionResultCellReportSupport.toCellBorderSideReport(side);
  }

  private static GridGrindResponse.WindowReport toWindowReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.Window window) {
    return new GridGrindResponse.WindowReport(
        window.sheetName(),
        window.topLeftAddress(),
        window.rowCount(),
        window.columnCount(),
        window.rows().stream()
            .map(
                row ->
                    new GridGrindResponse.WindowRowReport(
                        row.rowIndex(),
                        row.cells().stream().map(InspectionResultConverter::toCellReport).toList()))
            .toList());
  }

  private static GridGrindResponse.CellHyperlinkReport toCellHyperlinkReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.CellHyperlink hyperlink) {
    return new GridGrindResponse.CellHyperlinkReport(
        hyperlink.address(), toHyperlinkTarget(hyperlink.hyperlink()));
  }

  private static GridGrindResponse.CellCommentReport toCellCommentReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.CellComment comment) {
    return new GridGrindResponse.CellCommentReport(
        comment.address(), toCommentReport(comment.comment()));
  }

  private static GridGrindResponse.SheetLayoutReport toSheetLayoutReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout layout) {
    return new GridGrindResponse.SheetLayoutReport(
        layout.sheetName(),
        toPaneReport(layout.pane()),
        layout.zoomPercent(),
        toSheetPresentationReport(layout.presentation()),
        layout.columns().stream()
            .map(
                column ->
                    new GridGrindResponse.ColumnLayoutReport(
                        column.columnIndex(),
                        column.widthCharacters(),
                        column.hidden(),
                        column.outlineLevel(),
                        column.collapsed()))
            .toList(),
        layout.rows().stream()
            .map(
                row ->
                    new GridGrindResponse.RowLayoutReport(
                        row.rowIndex(),
                        row.heightPoints(),
                        row.hidden(),
                        row.outlineLevel(),
                        row.collapsed()))
            .toList());
  }

  private static PrintLayoutReport toPrintLayoutReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult printLayout) {
    return new PrintLayoutReport(
        printLayout.sheetName(),
        toPrintAreaReport(printLayout.printLayout().layout().printArea()),
        printLayout.printLayout().layout().orientation(),
        toPrintScalingReport(printLayout.printLayout().layout().scaling()),
        toPrintTitleRowsReport(printLayout.printLayout().layout().repeatingRows()),
        toPrintTitleColumnsReport(printLayout.printLayout().layout().repeatingColumns()),
        toHeaderFooterTextReport(printLayout.printLayout().layout().header()),
        toHeaderFooterTextReport(printLayout.printLayout().layout().footer()),
        new PrintSetupReport(
            new PrintMarginsReport(
                printLayout.printLayout().setup().margins().left(),
                printLayout.printLayout().setup().margins().right(),
                printLayout.printLayout().setup().margins().top(),
                printLayout.printLayout().setup().margins().bottom(),
                printLayout.printLayout().setup().margins().header(),
                printLayout.printLayout().setup().margins().footer()),
            printLayout.printLayout().setup().printGridlines(),
            printLayout.printLayout().setup().horizontallyCentered(),
            printLayout.printLayout().setup().verticallyCentered(),
            printLayout.printLayout().setup().paperSize(),
            printLayout.printLayout().setup().draft(),
            printLayout.printLayout().setup().blackAndWhite(),
            printLayout.printLayout().setup().copies(),
            printLayout.printLayout().setup().useFirstPageNumber(),
            printLayout.printLayout().setup().firstPageNumber(),
            printLayout.printLayout().setup().rowBreaks(),
            printLayout.printLayout().setup().columnBreaks()));
  }

  private static SheetPresentationReport toSheetPresentationReport(
      ExcelSheetPresentationSnapshot presentation) {
    return new SheetPresentationReport(
        new SheetDisplayReport(
            presentation.display().displayGridlines(),
            presentation.display().displayZeros(),
            presentation.display().displayRowColHeadings(),
            presentation.display().displayFormulas(),
            presentation.display().rightToLeft()),
        toCellColorReport(presentation.tabColor()),
        new SheetOutlineSummaryReport(
            presentation.outlineSummary().rowSumsBelow(),
            presentation.outlineSummary().rowSumsRight()),
        new SheetDefaultsReport(
            presentation.sheetDefaults().defaultColumnWidth(),
            presentation.sheetDefaults().defaultRowHeightPoints()),
        presentation.ignoredErrors().stream()
            .map(
                ignoredError ->
                    new IgnoredErrorReport(ignoredError.range(), ignoredError.errorTypes()))
            .toList());
  }

  private static PaneReport toPaneReport(ExcelSheetPane pane) {
    return switch (pane) {
      case ExcelSheetPane.None _ -> new PaneReport.None();
      case ExcelSheetPane.Frozen frozen ->
          new PaneReport.Frozen(
              frozen.splitColumn(), frozen.splitRow(), frozen.leftmostColumn(), frozen.topRow());
      case ExcelSheetPane.Split split ->
          new PaneReport.Split(
              split.xSplitPosition(),
              split.ySplitPosition(),
              split.leftmostColumn(),
              split.topRow(),
              split.activePane());
    };
  }

  private static PrintAreaReport toPrintAreaReport(ExcelPrintLayout.Area printArea) {
    return switch (printArea) {
      case ExcelPrintLayout.Area.None _ -> new PrintAreaReport.None();
      case ExcelPrintLayout.Area.Range range -> new PrintAreaReport.Range(range.range());
    };
  }

  private static PrintScalingReport toPrintScalingReport(ExcelPrintLayout.Scaling scaling) {
    return switch (scaling) {
      case ExcelPrintLayout.Scaling.Automatic _ -> new PrintScalingReport.Automatic();
      case ExcelPrintLayout.Scaling.Fit fit ->
          new PrintScalingReport.Fit(fit.widthPages(), fit.heightPages());
    };
  }

  private static PrintTitleRowsReport toPrintTitleRowsReport(
      ExcelPrintLayout.TitleRows repeatingRows) {
    return switch (repeatingRows) {
      case ExcelPrintLayout.TitleRows.None _ -> new PrintTitleRowsReport.None();
      case ExcelPrintLayout.TitleRows.Band band ->
          new PrintTitleRowsReport.Band(band.firstRowIndex(), band.lastRowIndex());
    };
  }

  private static PrintTitleColumnsReport toPrintTitleColumnsReport(
      ExcelPrintLayout.TitleColumns repeatingColumns) {
    return switch (repeatingColumns) {
      case ExcelPrintLayout.TitleColumns.None _ -> new PrintTitleColumnsReport.None();
      case ExcelPrintLayout.TitleColumns.Band band ->
          new PrintTitleColumnsReport.Band(band.firstColumnIndex(), band.lastColumnIndex());
    };
  }

  private static HeaderFooterTextReport toHeaderFooterTextReport(ExcelHeaderFooterText text) {
    return new HeaderFooterTextReport(text.left(), text.center(), text.right());
  }

  static GridGrindResponse.CellReport toCellReport(ExcelCellSnapshot snapshot) {
    return InspectionResultCellReportSupport.toCellReport(snapshot);
  }

  private static CellColorReport toCellColorReport(ExcelColorSnapshot color) {
    return InspectionResultCellReportSupport.toCellColorReport(color);
  }

  private static CustomXmlMappingReport toCustomXmlMappingReport(
      ExcelCustomXmlMappingSnapshot snapshot) {
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
            .map(InspectionResultConverter::toCustomXmlLinkedCellReport)
            .toList(),
        snapshot.linkedTables().stream()
            .map(InspectionResultConverter::toCustomXmlLinkedTableReport)
            .toList());
  }

  private static CustomXmlDataBindingReport toCustomXmlDataBindingReport(
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

  private static CustomXmlLinkedCellReport toCustomXmlLinkedCellReport(
      ExcelCustomXmlLinkedCellSnapshot snapshot) {
    return new CustomXmlLinkedCellReport(
        snapshot.sheetName(), snapshot.address(), snapshot.xpath(), snapshot.xmlDataType());
  }

  private static CustomXmlLinkedTableReport toCustomXmlLinkedTableReport(
      ExcelCustomXmlLinkedTableSnapshot snapshot) {
    return new CustomXmlLinkedTableReport(
        snapshot.sheetName(),
        snapshot.tableName(),
        snapshot.tableDisplayName(),
        snapshot.range(),
        snapshot.commonXPath());
  }

  private static CustomXmlExportReport toCustomXmlExportReport(
      ExcelCustomXmlExportSnapshot snapshot) {
    return new CustomXmlExportReport(
        toCustomXmlMappingReport(snapshot.mapping()),
        snapshot.encoding(),
        snapshot.schemaValidated(),
        snapshot.xml());
  }

  private static ArrayFormulaReport toArrayFormulaReport(ExcelArrayFormulaSnapshot snapshot) {
    return new ArrayFormulaReport(
        snapshot.sheetName(),
        snapshot.range(),
        snapshot.topLeftAddress(),
        snapshot.formula(),
        snapshot.singleCell());
  }

  private static GridGrindResponse.FormulaSurfaceReport toFormulaSurfaceReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurface analysis) {
    return InspectionResultAnalysisReportSupport.toFormulaSurfaceReport(analysis);
  }

  private static GridGrindResponse.SheetSchemaReport toSheetSchemaReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchema analysis) {
    return InspectionResultAnalysisReportSupport.toSheetSchemaReport(analysis);
  }

  private static GridGrindResponse.NamedRangeSurfaceReport toNamedRangeSurfaceReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurface analysis) {
    return InspectionResultAnalysisReportSupport.toNamedRangeSurfaceReport(analysis);
  }

  private static GridGrindResponse.FormulaHealthReport toFormulaHealthReport(
      WorkbookAnalysis.FormulaHealth analysis) {
    return InspectionResultAnalysisReportSupport.toFormulaHealthReport(analysis);
  }

  private static DataValidationHealthReport toDataValidationHealthReport(
      WorkbookAnalysis.DataValidationHealth analysis) {
    return InspectionResultAnalysisReportSupport.toDataValidationHealthReport(analysis);
  }

  private static ConditionalFormattingHealthReport toConditionalFormattingHealthReport(
      WorkbookAnalysis.ConditionalFormattingHealth analysis) {
    return InspectionResultAnalysisReportSupport.toConditionalFormattingHealthReport(analysis);
  }

  private static AutofilterHealthReport toAutofilterHealthReport(
      WorkbookAnalysis.AutofilterHealth analysis) {
    return InspectionResultAnalysisReportSupport.toAutofilterHealthReport(analysis);
  }

  private static TableHealthReport toTableHealthReport(WorkbookAnalysis.TableHealth analysis) {
    return InspectionResultAnalysisReportSupport.toTableHealthReport(analysis);
  }

  private static PivotTableHealthReport toPivotTableHealthReport(
      WorkbookAnalysis.PivotTableHealth analysis) {
    return InspectionResultAnalysisReportSupport.toPivotTableHealthReport(analysis);
  }

  private static GridGrindResponse.HyperlinkHealthReport toHyperlinkHealthReport(
      WorkbookAnalysis.HyperlinkHealth analysis) {
    return InspectionResultAnalysisReportSupport.toHyperlinkHealthReport(analysis);
  }

  private static GridGrindResponse.NamedRangeHealthReport toNamedRangeHealthReport(
      WorkbookAnalysis.NamedRangeHealth analysis) {
    return InspectionResultAnalysisReportSupport.toNamedRangeHealthReport(analysis);
  }

  private static GridGrindResponse.WorkbookFindingsReport toWorkbookFindingsReport(
      WorkbookAnalysis.WorkbookFindings analysis) {
    return InspectionResultAnalysisReportSupport.toWorkbookFindingsReport(analysis);
  }

  private static AutofilterEntryReport toAutofilterEntryReport(ExcelAutofilterSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelAutofilterSnapshot.SheetOwned sheetOwned ->
          new AutofilterEntryReport.SheetOwned(
              sheetOwned.range(),
              sheetOwned.filterColumns().stream()
                  .map(InspectionResultConverter::toAutofilterFilterColumnReport)
                  .toList(),
              toAutofilterSortStateReport(sheetOwned.sortState()));
      case ExcelAutofilterSnapshot.TableOwned tableOwned ->
          new AutofilterEntryReport.TableOwned(
              tableOwned.range(),
              tableOwned.tableName(),
              tableOwned.filterColumns().stream()
                  .map(InspectionResultConverter::toAutofilterFilterColumnReport)
                  .toList(),
              toAutofilterSortStateReport(tableOwned.sortState()));
    };
  }

  private static TableEntryReport toTableEntryReport(ExcelTableSnapshot snapshot) {
    return new TableEntryReport(
        snapshot.name(),
        snapshot.sheetName(),
        snapshot.range(),
        snapshot.headerRowCount(),
        snapshot.totalsRowCount(),
        snapshot.columnNames(),
        snapshot.columns().stream().map(InspectionResultConverter::toTableColumnReport).toList(),
        toTableStyleReport(snapshot.style()),
        snapshot.hasAutofilter(),
        snapshot.comment(),
        snapshot.published(),
        snapshot.insertRow(),
        snapshot.insertRowShift(),
        snapshot.headerRowCellStyle(),
        snapshot.dataCellStyle(),
        snapshot.totalsRowCellStyle());
  }

  private static PivotTableReport toPivotTableReport(ExcelPivotTableSnapshot snapshot) {
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

  private static GridGrindResponse.CommentReport toCommentReport(ExcelCommentSnapshot comment) {
    return InspectionResultCellReportSupport.toCommentReport(comment);
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
    return sortState == null
        ? null
        : new AutofilterSortStateReport(
            sortState.range(),
            sortState.caseSensitive(),
            sortState.columnSort(),
            sortState.sortMethod(),
            sortState.conditions().stream()
                .map(
                    condition ->
                        new AutofilterSortConditionReport(
                            condition.range(),
                            condition.descending(),
                            condition.sortBy(),
                            toCellColorReport(condition.color()),
                            condition.iconId()))
                .toList());
  }

  private static TableColumnReport toTableColumnReport(ExcelTableColumnSnapshot column) {
    return new TableColumnReport(
        column.id(),
        column.name(),
        column.uniqueName(),
        column.totalsRowLabel(),
        column.totalsRowFunction(),
        column.calculatedColumnFormula());
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

  private static SheetProtectionSettings toSheetProtectionSettings(
      ExcelSheetProtectionSettings settings) {
    return new SheetProtectionSettings(
        settings.autoFilterLocked(),
        settings.deleteColumnsLocked(),
        settings.deleteRowsLocked(),
        settings.formatCellsLocked(),
        settings.formatColumnsLocked(),
        settings.formatRowsLocked(),
        settings.insertColumnsLocked(),
        settings.insertHyperlinksLocked(),
        settings.insertRowsLocked(),
        settings.objectsLocked(),
        settings.pivotTablesLocked(),
        settings.scenariosLocked(),
        settings.selectLockedCellsLocked(),
        settings.selectUnlockedCellsLocked(),
        settings.sortLocked());
  }
}
