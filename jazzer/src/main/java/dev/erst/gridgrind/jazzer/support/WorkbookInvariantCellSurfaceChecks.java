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
import dev.erst.gridgrind.contract.dto.CellGradientStopReport;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.CommentAnchorReport;
import dev.erst.gridgrind.contract.dto.CustomXmlDataBindingReport;
import dev.erst.gridgrind.contract.dto.CustomXmlExportReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingReport;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.PrintMarginsReport;
import dev.erst.gridgrind.contract.dto.PrintSetupReport;
import dev.erst.gridgrind.contract.dto.RichTextRunReport;
import dev.erst.gridgrind.contract.dto.TableColumnReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;

/** Owns cell-surface invariant checks for reports, styles, security, and XML-linked metadata. */
final class WorkbookInvariantCellSurfaceChecks {
  private WorkbookInvariantCellSurfaceChecks() {}

  static void requireCellReportShape(GridGrindResponse.CellReport cellReport) {
    WorkbookInvariantChecks.require(cellReport.address() != null, "cell address must not be null");
    WorkbookInvariantChecks.require(
        !cellReport.address().isBlank(), "cell address must not be blank");
    WorkbookInvariantChecks.require(
        cellReport.declaredType() != null, "declaredType must not be null");
    WorkbookInvariantChecks.require(
        cellReport.effectiveType() != null, "effectiveType must not be null");
    WorkbookInvariantChecks.require(
        cellReport.displayValue() != null, "displayValue must not be null");
    requireCellStyleShape(cellReport.style());

    switch (cellReport) {
      case GridGrindResponse.CellReport.BlankReport _ -> {}
      case GridGrindResponse.CellReport.TextReport text -> {
        WorkbookInvariantChecks.require(text.stringValue() != null, "stringValue must not be null");
        if (text.richText() != null) {
          WorkbookInvariantChecks.require(!text.richText().isEmpty(), "richText must not be empty");
          StringBuilder builder = new StringBuilder();
          for (var run : text.richText()) {
            WorkbookInvariantChecks.require(
                run.text() != null, "richText run text must not be null");
            WorkbookInvariantChecks.require(
                !run.text().isEmpty(), "richText run text must not be empty");
            requireCellFontShape(run.font());
            builder.append(run.text());
          }
          WorkbookInvariantChecks.require(
              text.stringValue().equals(builder.toString()),
              "richText run text must concatenate to stringValue");
        }
      }
      case GridGrindResponse.CellReport.NumberReport number ->
          WorkbookInvariantChecks.require(
              number.numberValue() != null, "numberValue must not be null");
      case GridGrindResponse.CellReport.BooleanReport bool ->
          WorkbookInvariantChecks.require(
              bool.booleanValue() != null, "booleanValue must not be null");
      case GridGrindResponse.CellReport.ErrorReport error ->
          WorkbookInvariantChecks.require(
              error.errorValue() != null, "errorValue must not be null");
      case GridGrindResponse.CellReport.FormulaReport formula -> {
        WorkbookInvariantChecks.require(formula.formula() != null, "formula must not be null");
        requireCellReportShape(formula.evaluation());
      }
    }
    if (cellReport.hyperlink() != null) {
      requireHyperlinkShape(cellReport.hyperlink());
    }
    if (cellReport.comment() != null) {
      requireCommentReportShape(cellReport.comment());
    }
  }

  static void requireCommentReportShape(GridGrindResponse.CommentReport comment) {
    WorkbookInvariantChecks.require(comment.text() != null, "comment text must not be null");
    WorkbookInvariantChecks.require(comment.author() != null, "comment author must not be null");
    WorkbookInvariantChecks.require(!comment.text().isBlank(), "comment text must not be blank");
    WorkbookInvariantChecks.require(
        !comment.author().isBlank(), "comment author must not be blank");
    if (comment.runs() != null) {
      WorkbookInvariantChecks.require(!comment.runs().isEmpty(), "comment runs must not be empty");
      StringBuilder builder = new StringBuilder();
      for (RichTextRunReport run : comment.runs()) {
        WorkbookInvariantChecks.require(run != null, "comment runs must not contain null values");
        WorkbookInvariantChecks.require(run.text() != null, "comment run text must not be null");
        WorkbookInvariantChecks.require(
            !run.text().isEmpty(), "comment run text must not be empty");
        requireCellFontShape(run.font());
        builder.append(run.text());
      }
      WorkbookInvariantChecks.require(
          builder.toString().equals(comment.text()), "comment runs must concatenate to text");
    }
    if (comment.anchor() != null) {
      requireCommentAnchorShape(comment.anchor());
    }
  }

  static void requireNamedRangeShape(GridGrindResponse.NamedRangeReport namedRange) {
    WorkbookInvariantChecks.require(namedRange.name() != null, "namedRange name must not be null");
    WorkbookInvariantChecks.require(
        !namedRange.name().isBlank(), "namedRange name must not be blank");
    WorkbookInvariantChecks.require(
        namedRange.scope() != null, "namedRange scope must not be null");
    WorkbookInvariantChecks.require(
        namedRange.refersToFormula() != null, "namedRange formula must not be null");

    switch (namedRange) {
      case GridGrindResponse.NamedRangeReport.RangeReport range -> {
        WorkbookInvariantChecks.require(
            range.target() != null, "namedRange target must not be null");
        WorkbookInvariantChecks.require(
            range.target().sheetName() != null, "namedRange target sheet must not be null");
        WorkbookInvariantChecks.require(
            range.target().range() != null, "namedRange target range must not be null");
        WorkbookInvariantChecks.require(
            !range.target().sheetName().isBlank(), "namedRange target sheet must not be blank");
        WorkbookInvariantChecks.require(
            !range.target().range().isBlank(), "namedRange target range must not be blank");
      }
      case GridGrindResponse.NamedRangeReport.FormulaReport _ -> {}
    }
  }

  static void requireHyperlinkShape(HyperlinkTarget hyperlink) {
    WorkbookInvariantChecks.require(hyperlink != null, "hyperlink must not be null");
    switch (hyperlink) {
      case HyperlinkTarget.Url url -> {
        WorkbookInvariantChecks.require(url.target() != null, "hyperlink target must not be null");
        WorkbookInvariantChecks.require(
            !url.target().isBlank(), "hyperlink target must not be blank");
        WorkbookInvariantChecks.require(
            !url.target().regionMatches(true, 0, "file:", 0, 5),
            "URL hyperlink targets must not use file: schemes");
        WorkbookInvariantChecks.require(
            !url.target().regionMatches(true, 0, "mailto:", 0, 7),
            "URL hyperlink targets must not use mailto: schemes");
      }
      case HyperlinkTarget.Email email -> {
        WorkbookInvariantChecks.require(email.email() != null, "hyperlink email must not be null");
        WorkbookInvariantChecks.require(
            !email.email().isBlank(), "hyperlink email must not be blank");
        WorkbookInvariantChecks.require(
            !email.email().regionMatches(true, 0, "mailto:", 0, 7),
            "EMAIL hyperlink targets must omit the mailto: prefix");
      }
      case HyperlinkTarget.File file -> {
        WorkbookInvariantChecks.require(file.path() != null, "hyperlink path must not be null");
        WorkbookInvariantChecks.require(!file.path().isBlank(), "hyperlink path must not be blank");
        WorkbookInvariantChecks.require(
            !file.path().regionMatches(true, 0, "file:", 0, 5),
            "FILE hyperlink targets must be normalized path strings");
      }
      case HyperlinkTarget.Document document -> {
        WorkbookInvariantChecks.require(
            document.target() != null, "hyperlink target must not be null");
        WorkbookInvariantChecks.require(
            !document.target().isBlank(), "hyperlink target must not be blank");
      }
    }
  }

  static void requireCellStyleShape(GridGrindResponse.CellStyleReport style) {
    WorkbookInvariantChecks.require(style != null, "style must not be null");
    WorkbookInvariantChecks.require(style.numberFormat() != null, "numberFormat must not be null");
    requireCellAlignmentShape(style.alignment());
    requireCellFontShape(style.font());
    requireCellFillShape(style.fill());
    requireCellBorderShape(style.border());
    requireCellProtectionShape(style.protection());
  }

  static void requireCellAlignmentShape(CellAlignmentReport alignment) {
    WorkbookInvariantChecks.require(alignment != null, "alignment must not be null");
    WorkbookInvariantChecks.require(
        alignment.horizontalAlignment() != null, "horizontalAlignment must not be null");
    WorkbookInvariantChecks.require(
        alignment.verticalAlignment() != null, "verticalAlignment must not be null");
    WorkbookInvariantChecks.require(
        alignment.textRotation() >= 0 && alignment.textRotation() <= 180,
        "textRotation must be between 0 and 180 inclusive");
    WorkbookInvariantChecks.require(
        alignment.indentation() >= 0 && alignment.indentation() <= 250,
        "indentation must be between 0 and 250 inclusive");
  }

  static void requireCellFontShape(CellFontReport font) {
    WorkbookInvariantChecks.require(font != null, "font must not be null");
    WorkbookInvariantChecks.require(font.fontName() != null, "fontName must not be null");
    WorkbookInvariantChecks.require(!font.fontName().isBlank(), "fontName must not be blank");
    requireFontHeightShape(font.fontHeight());
    if (font.fontColor() != null) {
      requireCellColorShape(font.fontColor(), "fontColor");
    }
  }

  static void requireCellFillShape(CellFillReport fill) {
    WorkbookInvariantChecks.require(fill != null, "fill must not be null");
    WorkbookInvariantChecks.require(fill.pattern() != null, "fill pattern must not be null");
    if (fill.foregroundColor() != null) {
      requireCellColorShape(fill.foregroundColor(), "fill foregroundColor");
    }
    if (fill.backgroundColor() != null) {
      requireCellColorShape(fill.backgroundColor(), "fill backgroundColor");
    }
    if (fill.gradient() != null) {
      requireCellGradientFillShape(fill.gradient());
      WorkbookInvariantChecks.require(
          fill.foregroundColor() == null && fill.backgroundColor() == null,
          "gradient fills must not carry flat colors");
    }
    if (fill.pattern() == ExcelFillPattern.NONE && fill.gradient() == null) {
      WorkbookInvariantChecks.require(
          fill.foregroundColor() == null && fill.backgroundColor() == null,
          "fill pattern NONE must not carry colors");
    }
    if (fill.pattern() == ExcelFillPattern.SOLID && fill.gradient() == null) {
      WorkbookInvariantChecks.require(
          fill.backgroundColor() == null, "SOLID fills must not carry backgroundColor");
    }
  }

  static void requireCellBorderShape(CellBorderReport border) {
    WorkbookInvariantChecks.require(border != null, "border must not be null");
    requireCellBorderSideShape(border.top(), "top");
    requireCellBorderSideShape(border.right(), "right");
    requireCellBorderSideShape(border.bottom(), "bottom");
    requireCellBorderSideShape(border.left(), "left");
  }

  static void requireCellBorderSideShape(CellBorderSideReport side, String label) {
    WorkbookInvariantChecks.require(side != null, label + " border side must not be null");
    WorkbookInvariantChecks.require(side.style() != null, label + " border style must not be null");
    if (side.color() != null) {
      requireCellColorShape(side.color(), label + " border color");
    }
  }

  static void requireWorkbookProtectionShape(WorkbookProtectionReport protection) {
    WorkbookInvariantChecks.require(protection != null, "workbook protection must not be null");
  }

  static void requirePackageSecurityShape(OoxmlPackageSecurityReport security) {
    WorkbookInvariantChecks.require(security != null, "package security must not be null");
    WorkbookInvariantChecks.require(
        security.encryption() != null, "package encryption must not be null");
    if (security.encryption().encrypted()) {
      WorkbookInvariantChecks.require(
          security.encryption().mode() != null, "encrypted package mode must not be null");
      WorkbookInvariantChecks.requireNonBlank(
          security.encryption().cipherAlgorithm(), "encrypted package cipherAlgorithm");
      WorkbookInvariantChecks.requireNonBlank(
          security.encryption().hashAlgorithm(), "encrypted package hashAlgorithm");
      WorkbookInvariantChecks.requireNonBlank(
          security.encryption().chainingMode(), "encrypted package chainingMode");
      WorkbookInvariantChecks.require(
          security.encryption().keyBits() != null && security.encryption().keyBits() > 0,
          "encrypted package keyBits must be positive");
      WorkbookInvariantChecks.require(
          security.encryption().blockSize() != null && security.encryption().blockSize() > 0,
          "encrypted package blockSize must be positive");
      WorkbookInvariantChecks.require(
          security.encryption().spinCount() != null && security.encryption().spinCount() >= 0,
          "encrypted package spinCount must be zero or positive");
    }
    WorkbookInvariantChecks.require(
        security.signatures() != null, "package signatures must not be null");
    security
        .signatures()
        .forEach(
            signature -> {
              WorkbookInvariantChecks.require(
                  signature != null, "package signature must not be null");
              WorkbookInvariantChecks.requireNonBlank(
                  signature.packagePartName(), "package signature part");
              WorkbookInvariantChecks.require(
                  signature.state() != null, "package signature state must not be null");
            });
  }

  static void requireCustomXmlMappingShape(CustomXmlMappingReport mapping) {
    WorkbookInvariantChecks.require(mapping != null, "custom XML mapping must not be null");
    WorkbookInvariantChecks.require(
        mapping.mapId() > 0L, "custom XML mapping mapId must be positive");
    WorkbookInvariantChecks.requireNonBlank(mapping.name(), "custom XML mapping name");
    WorkbookInvariantChecks.requireNonBlank(
        mapping.rootElement(), "custom XML mapping rootElement");
    WorkbookInvariantChecks.requireNonBlank(mapping.schemaId(), "custom XML mapping schemaId");
    if (mapping.schemaNamespace() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          mapping.schemaNamespace(), "custom XML mapping schemaNamespace");
    }
    if (mapping.schemaLanguage() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          mapping.schemaLanguage(), "custom XML mapping schemaLanguage");
    }
    if (mapping.schemaReference() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          mapping.schemaReference(), "custom XML mapping schemaReference");
    }
    if (mapping.schemaXml() != null) {
      WorkbookInvariantChecks.requireNonBlank(mapping.schemaXml(), "custom XML mapping schemaXml");
    }
    if (mapping.dataBinding() != null) {
      requireCustomXmlDataBindingShape(mapping.dataBinding());
    }
    mapping
        .linkedCells()
        .forEach(WorkbookInvariantCellSurfaceChecks::requireCustomXmlLinkedCellShape);
    mapping
        .linkedTables()
        .forEach(WorkbookInvariantCellSurfaceChecks::requireCustomXmlLinkedTableShape);
  }

  static void requireCustomXmlDataBindingShape(CustomXmlDataBindingReport dataBinding) {
    WorkbookInvariantChecks.require(
        dataBinding != null, "custom XML data binding must not be null");
    if (dataBinding.dataBindingName() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          dataBinding.dataBindingName(), "custom XML dataBindingName");
    }
    if (dataBinding.connectionId() != null) {
      WorkbookInvariantChecks.require(
          dataBinding.connectionId() >= 0L, "custom XML connectionId must not be negative");
    }
    if (dataBinding.fileBindingName() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          dataBinding.fileBindingName(), "custom XML fileBindingName");
    }
    WorkbookInvariantChecks.require(
        dataBinding.loadMode() >= 0L, "custom XML loadMode must not be negative");
  }

  static void requireCustomXmlLinkedCellShape(CustomXmlLinkedCellReport linkedCell) {
    WorkbookInvariantChecks.require(linkedCell != null, "custom XML linked cell must not be null");
    WorkbookInvariantChecks.requireNonBlank(
        linkedCell.sheetName(), "custom XML linked cell sheetName");
    WorkbookInvariantChecks.requireNonBlank(linkedCell.address(), "custom XML linked cell address");
    WorkbookInvariantChecks.requireNonBlank(linkedCell.xpath(), "custom XML linked cell xpath");
    WorkbookInvariantChecks.requireNonBlank(
        linkedCell.xmlDataType(), "custom XML linked cell xmlDataType");
  }

  static void requireCustomXmlLinkedTableShape(CustomXmlLinkedTableReport linkedTable) {
    WorkbookInvariantChecks.require(
        linkedTable != null, "custom XML linked table must not be null");
    WorkbookInvariantChecks.requireNonBlank(
        linkedTable.sheetName(), "custom XML linked table sheetName");
    WorkbookInvariantChecks.requireNonBlank(
        linkedTable.tableName(), "custom XML linked table tableName");
    WorkbookInvariantChecks.requireNonBlank(
        linkedTable.tableDisplayName(), "custom XML linked table tableDisplayName");
    WorkbookInvariantChecks.requireNonBlank(linkedTable.range(), "custom XML linked table range");
    WorkbookInvariantChecks.requireNonBlank(
        linkedTable.commonXPath(), "custom XML linked table commonXPath");
  }

  static void requireCustomXmlExportShape(CustomXmlExportReport export) {
    WorkbookInvariantChecks.require(export != null, "custom XML export must not be null");
    requireCustomXmlMappingShape(export.mapping());
    WorkbookInvariantChecks.requireNonBlank(export.encoding(), "custom XML export encoding");
    WorkbookInvariantChecks.requireNonBlank(export.xml(), "custom XML export xml");
  }

  static void requireArrayFormulaShape(ArrayFormulaReport arrayFormula) {
    WorkbookInvariantChecks.require(arrayFormula != null, "array formula must not be null");
    WorkbookInvariantChecks.requireNonBlank(arrayFormula.sheetName(), "array formula sheetName");
    WorkbookInvariantChecks.requireNonBlank(arrayFormula.range(), "array formula range");
    WorkbookInvariantChecks.requireNonBlank(
        arrayFormula.topLeftAddress(), "array formula topLeftAddress");
    WorkbookInvariantChecks.requireNonBlank(arrayFormula.formula(), "array formula formula");
  }

  static void requireCommentAnchorShape(CommentAnchorReport anchor) {
    WorkbookInvariantChecks.require(
        anchor.firstColumn() >= 0, "comment anchor firstColumn must not be negative");
    WorkbookInvariantChecks.require(
        anchor.firstRow() >= 0, "comment anchor firstRow must not be negative");
    WorkbookInvariantChecks.require(
        anchor.lastColumn() >= anchor.firstColumn(), "comment anchor columns must be ordered");
    WorkbookInvariantChecks.require(
        anchor.lastRow() >= anchor.firstRow(), "comment anchor rows must be ordered");
  }

  static void requirePrintSetupShape(PrintSetupReport setup) {
    WorkbookInvariantChecks.require(setup != null, "print setup must not be null");
    requirePrintMarginsShape(setup.margins());
    WorkbookInvariantChecks.require(
        setup.paperSize() >= 0, "print setup paperSize must not be negative");
    WorkbookInvariantChecks.require(setup.copies() >= 0, "print setup copies must not be negative");
    WorkbookInvariantChecks.require(
        setup.firstPageNumber() >= 0, "print setup firstPageNumber must not be negative");
    WorkbookInvariantChecks.require(
        setup.rowBreaks() != null, "print setup rowBreaks must not be null");
    WorkbookInvariantChecks.require(
        setup.columnBreaks() != null, "print setup columnBreaks must not be null");
    setup
        .rowBreaks()
        .forEach(
            rowBreak ->
                WorkbookInvariantChecks.require(
                    rowBreak >= 0, "print setup rowBreak must not be negative"));
    setup
        .columnBreaks()
        .forEach(
            columnBreak ->
                WorkbookInvariantChecks.require(
                    columnBreak >= 0, "print setup columnBreak must not be negative"));
  }

  static void requirePrintMarginsShape(PrintMarginsReport margins) {
    WorkbookInvariantChecks.require(margins != null, "print margins must not be null");
  }

  static void requireAutofilterFilterColumnShape(AutofilterFilterColumnReport filterColumn) {
    WorkbookInvariantChecks.require(
        filterColumn != null, "autofilter filterColumn must not be null");
    WorkbookInvariantChecks.require(
        filterColumn.columnId() >= 0L, "autofilter columnId must not be negative");
    requireAutofilterCriterionShape(filterColumn.criterion());
  }

  static void requireAutofilterCriterionShape(AutofilterFilterCriterionReport criterion) {
    WorkbookInvariantChecks.require(criterion != null, "autofilter criterion must not be null");
    switch (criterion) {
      case AutofilterFilterCriterionReport.Values values -> {
        WorkbookInvariantChecks.require(
            values.values() != null, "autofilter values must not be null");
        values
            .values()
            .forEach(
                value ->
                    WorkbookInvariantChecks.require(
                        value != null, "autofilter value must not be null"));
      }
      case AutofilterFilterCriterionReport.Custom custom -> {
        WorkbookInvariantChecks.require(
            custom.conditions() != null, "autofilter custom conditions must not be null");
        WorkbookInvariantChecks.require(
            !custom.conditions().isEmpty(), "autofilter custom conditions must not be empty");
        custom
            .conditions()
            .forEach(
                condition -> {
                  WorkbookInvariantChecks.require(
                      condition != null, "autofilter custom condition must not be null");
                  WorkbookInvariantChecks.requireNonBlank(
                      condition.operator(), "autofilter custom operator");
                  WorkbookInvariantChecks.requireNonBlank(
                      condition.value(), "autofilter custom value");
                });
      }
      case AutofilterFilterCriterionReport.Dynamic dynamic -> {
        WorkbookInvariantChecks.requireNonBlank(dynamic.type(), "autofilter dynamic type");
        if (dynamic.value() != null) {
          WorkbookInvariantChecks.require(
              Double.isFinite(dynamic.value()), "autofilter dynamic value must be finite");
        }
        if (dynamic.maxValue() != null) {
          WorkbookInvariantChecks.require(
              Double.isFinite(dynamic.maxValue()), "autofilter dynamic maxValue must be finite");
        }
      }
      case AutofilterFilterCriterionReport.Top10 top10 -> {
        WorkbookInvariantChecks.require(
            Double.isFinite(top10.value()), "autofilter top10 value must be finite");
        WorkbookInvariantChecks.require(
            top10.value() >= 0.0d, "autofilter top10 value must not be negative");
        if (top10.filterValue() != null) {
          WorkbookInvariantChecks.require(
              Double.isFinite(top10.filterValue()), "autofilter top10 filterValue must be finite");
        }
      }
      case AutofilterFilterCriterionReport.Color color -> {
        if (color.color() != null) {
          requireCellColorShape(color.color(), "autofilter color");
        }
      }
      case AutofilterFilterCriterionReport.Icon icon -> {
        WorkbookInvariantChecks.requireNonBlank(icon.iconSet(), "autofilter iconSet");
        WorkbookInvariantChecks.require(
            icon.iconId() >= 0, "autofilter iconId must not be negative");
      }
    }
  }

  static void requireAutofilterSortStateShape(AutofilterSortStateReport sortState) {
    WorkbookInvariantChecks.require(sortState != null, "autofilter sortState must not be null");
    WorkbookInvariantChecks.requireNonBlank(sortState.range(), "autofilter sortState range");
    WorkbookInvariantChecks.require(
        sortState.conditions() != null, "autofilter sortState conditions must not be null");
    sortState
        .conditions()
        .forEach(WorkbookInvariantCellSurfaceChecks::requireAutofilterSortConditionShape);
  }

  static void requireAutofilterSortConditionShape(AutofilterSortConditionReport condition) {
    WorkbookInvariantChecks.require(
        condition != null, "autofilter sort condition must not be null");
    WorkbookInvariantChecks.requireNonBlank(condition.range(), "autofilter sort condition range");
    if (condition.color() != null) {
      requireCellColorShape(condition.color(), "autofilter sort color");
    }
    if (condition.iconId() != null) {
      WorkbookInvariantChecks.require(
          condition.iconId() >= 0, "autofilter sort iconId must not be negative");
    }
  }

  static void requireTableColumnShape(TableColumnReport column) {
    WorkbookInvariantChecks.require(column != null, "table column must not be null");
    WorkbookInvariantChecks.require(column.id() >= 0L, "table column id must not be negative");
    WorkbookInvariantChecks.require(column.name() != null, "table column name must not be null");
  }

  static void requireCellGradientFillShape(CellGradientFillReport gradient) {
    WorkbookInvariantChecks.require(gradient != null, "gradient fill must not be null");
    WorkbookInvariantChecks.requireNonBlank(gradient.type(), "gradient fill type");
    WorkbookInvariantChecks.require(
        gradient.stops() != null, "gradient fill stops must not be null");
    WorkbookInvariantChecks.require(
        !gradient.stops().isEmpty(), "gradient fill stops must not be empty");
    for (CellGradientStopReport stop : gradient.stops()) {
      WorkbookInvariantChecks.require(stop != null, "gradient fill stop must not be null");
      WorkbookInvariantChecks.require(
          Double.isFinite(stop.position()) && stop.position() >= 0.0d && stop.position() <= 1.0d,
          "gradient fill stop position must be between 0.0 and 1.0");
      requireCellColorShape(stop.color(), "gradient fill stop color");
    }
  }

  static void requireCellColorShape(CellColorReport color, String label) {
    WorkbookInvariantChecks.require(color != null, label + " must not be null");
    WorkbookInvariantChecks.require(
        color.rgb() != null || color.theme() != null || color.indexed() != null,
        label + " must expose rgb, theme, or indexed semantics");
    if (color.rgb() != null) {
      WorkbookInvariantChecks.requireNonBlank(color.rgb(), label + " rgb");
    }
    if (color.theme() != null) {
      WorkbookInvariantChecks.require(color.theme() >= 0, label + " theme must not be negative");
    }
    if (color.indexed() != null) {
      WorkbookInvariantChecks.require(
          color.indexed() >= 0, label + " indexed must not be negative");
    }
    if (color.tint() != null) {
      WorkbookInvariantChecks.require(
          Double.isFinite(color.tint()), label + " tint must be finite");
    }
  }

  static void requireCellProtectionShape(CellProtectionReport protection) {
    WorkbookInvariantChecks.require(protection != null, "protection must not be null");
  }

  static void requireFontHeightShape(FontHeightReport fontHeight) {
    WorkbookInvariantChecks.require(fontHeight != null, "fontHeight must not be null");
    ExcelFontHeight expected = new ExcelFontHeight(fontHeight.twips());
    WorkbookInvariantChecks.require(
        expected.points().compareTo(fontHeight.points()) == 0,
        "fontHeight points must match twips");
  }
}
