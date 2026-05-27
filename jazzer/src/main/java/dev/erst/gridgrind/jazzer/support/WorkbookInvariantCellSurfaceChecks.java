package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.ArrayFormulaReport;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnReport;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionReport;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateReport;
import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellGradientFillReport;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.CellStyleReport;
import dev.erst.gridgrind.contract.dto.CommentAnchorReport;
import dev.erst.gridgrind.contract.dto.CustomXmlDataBindingReport;
import dev.erst.gridgrind.contract.dto.CustomXmlExportReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingReport;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.PrintMarginsReport;
import dev.erst.gridgrind.contract.dto.PrintSetupReport;
import dev.erst.gridgrind.contract.dto.TableColumnReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;

/** Owns cell-surface invariant checks for reports, styles, security, and XML-linked metadata. */
final class WorkbookInvariantCellSurfaceChecks {
  private WorkbookInvariantCellSurfaceChecks() {}

  static void requireCellReportShape(dev.erst.gridgrind.contract.dto.CellReport cellReport) {
    WorkbookInvariantCellContentChecks.requireCellReportShape(cellReport);
  }

  static void requireCommentReportShape(dev.erst.gridgrind.contract.dto.CommentReport comment) {
    WorkbookInvariantCellContentChecks.requireCommentReportShape(comment);
  }

  static void requireNamedRangeShape(dev.erst.gridgrind.contract.dto.NamedRangeReport namedRange) {
    WorkbookInvariantCellContentChecks.requireNamedRangeShape(namedRange);
  }

  static void requireHyperlinkShape(dev.erst.gridgrind.contract.dto.HyperlinkTarget hyperlink) {
    WorkbookInvariantCellContentChecks.requireHyperlinkShape(hyperlink);
  }

  static void requireCellStyleShape(CellStyleReport style) {
    WorkbookInvariantCellStyleChecks.requireCellStyleShape(style);
  }

  static void requireCellAlignmentShape(CellAlignmentReport alignment) {
    WorkbookInvariantCellStyleChecks.requireCellAlignmentShape(alignment);
  }

  static void requireCellFontShape(CellFontReport font) {
    WorkbookInvariantCellStyleChecks.requireCellFontShape(font);
  }

  static void requireCellFillShape(CellFillReport fill) {
    WorkbookInvariantCellStyleChecks.requireCellFillShape(fill);
  }

  static void requireCellBorderShape(CellBorderReport border) {
    WorkbookInvariantCellStyleChecks.requireCellBorderShape(border);
  }

  static void requireCellBorderSideShape(CellBorderSideReport side, String label) {
    WorkbookInvariantCellStyleChecks.requireCellBorderSideShape(side, label);
  }

  static void requireWorkbookProtectionShape(WorkbookProtectionReport protection) {
    WorkbookInvariantCellMetadataChecks.requireWorkbookProtectionShape(protection);
  }

  static void requirePackageSecurityShape(OoxmlPackageSecurityReport security) {
    WorkbookInvariantCellMetadataChecks.requirePackageSecurityShape(security);
  }

  static void requireCustomXmlMappingShape(CustomXmlMappingReport mapping) {
    WorkbookInvariantCellMetadataChecks.requireCustomXmlMappingShape(mapping);
  }

  static void requireCustomXmlDataBindingShape(CustomXmlDataBindingReport dataBinding) {
    WorkbookInvariantCellMetadataChecks.requireCustomXmlDataBindingShape(dataBinding);
  }

  static void requireCustomXmlLinkedCellShape(CustomXmlLinkedCellReport linkedCell) {
    WorkbookInvariantCellMetadataChecks.requireCustomXmlLinkedCellShape(linkedCell);
  }

  static void requireCustomXmlLinkedTableShape(CustomXmlLinkedTableReport linkedTable) {
    WorkbookInvariantCellMetadataChecks.requireCustomXmlLinkedTableShape(linkedTable);
  }

  static void requireCustomXmlExportShape(CustomXmlExportReport export) {
    WorkbookInvariantCellMetadataChecks.requireCustomXmlExportShape(export);
  }

  static void requireArrayFormulaShape(ArrayFormulaReport arrayFormula) {
    WorkbookInvariantCellMetadataChecks.requireArrayFormulaShape(arrayFormula);
  }

  static void requireCommentAnchorShape(CommentAnchorReport anchor) {
    WorkbookInvariantCellMetadataChecks.requireCommentAnchorShape(anchor);
  }

  static void requirePrintSetupShape(PrintSetupReport setup) {
    WorkbookInvariantCellMetadataChecks.requirePrintSetupShape(setup);
  }

  static void requirePrintMarginsShape(PrintMarginsReport margins) {
    WorkbookInvariantCellMetadataChecks.requirePrintMarginsShape(margins);
  }

  static void requireAutofilterFilterColumnShape(AutofilterFilterColumnReport filterColumn) {
    WorkbookInvariantCellMetadataChecks.requireAutofilterFilterColumnShape(filterColumn);
  }

  static void requireAutofilterCriterionShape(AutofilterFilterCriterionReport criterion) {
    WorkbookInvariantCellMetadataChecks.requireAutofilterCriterionShape(criterion);
  }

  static void requireAutofilterSortStateShape(AutofilterSortStateReport sortState) {
    WorkbookInvariantCellMetadataChecks.requireAutofilterSortStateShape(sortState);
  }

  static void requireAutofilterSortConditionShape(AutofilterSortConditionReport condition) {
    WorkbookInvariantCellMetadataChecks.requireAutofilterSortConditionShape(condition);
  }

  static void requireTableColumnShape(TableColumnReport column) {
    WorkbookInvariantCellMetadataChecks.requireTableColumnShape(column);
  }

  static void requireCellGradientFillShape(CellGradientFillReport gradient) {
    WorkbookInvariantCellStyleChecks.requireCellGradientFillShape(gradient);
  }

  static void requireCellColorShape(CellColorReport color, String label) {
    WorkbookInvariantCellStyleChecks.requireCellColorShape(color, label);
  }

  static void requireCellProtectionShape(CellProtectionReport protection) {
    WorkbookInvariantCellStyleChecks.requireCellProtectionShape(protection);
  }

  static void requireFontHeightShape(FontHeightReport fontHeight) {
    WorkbookInvariantCellStyleChecks.requireFontHeightShape(fontHeight);
  }
}
