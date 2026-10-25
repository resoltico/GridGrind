package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.core.model.customxml.CustomXmlDataBindingReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlExportReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.model.drawing.CommentAnchorReport;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.print.PrintMarginsReport;
import dev.erst.gridgrind.core.model.print.PrintSetupReport;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellGradientFillReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.table.TableColumnReport;

/** Owns cell-surface invariant checks for reports, styles, security, and XML-linked metadata. */
final class WorkbookInvariantCellSurfaceChecks {
  private WorkbookInvariantCellSurfaceChecks() {}

  static void requireCellReportShape(dev.erst.gridgrind.core.model.value.CellReport cellReport) {
    WorkbookInvariantCellContentChecks.requireCellReportShape(cellReport);
  }

  static void requireCommentReportShape(
      dev.erst.gridgrind.core.model.drawing.CommentReport comment) {
    WorkbookInvariantCellContentChecks.requireCommentReportShape(comment);
  }

  static void requireNamedRangeShape(
      dev.erst.gridgrind.core.model.formula.NamedRangeReport namedRange) {
    WorkbookInvariantCellContentChecks.requireNamedRangeShape(namedRange);
  }

  static void requireHyperlinkShape(dev.erst.gridgrind.core.model.value.HyperlinkTarget hyperlink) {
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
