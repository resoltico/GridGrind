package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.ArrayFormulaReport;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnReport;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionReport;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateReport;
import dev.erst.gridgrind.contract.dto.CommentAnchorReport;
import dev.erst.gridgrind.contract.dto.CustomXmlDataBindingReport;
import dev.erst.gridgrind.contract.dto.CustomXmlExportReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingReport;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.PrintMarginsReport;
import dev.erst.gridgrind.contract.dto.PrintSetupReport;
import dev.erst.gridgrind.contract.dto.TableColumnReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;

/** Owns invariant checks for metadata-oriented cell and workbook report payloads. */
final class WorkbookInvariantCellMetadataChecks {
  private WorkbookInvariantCellMetadataChecks() {}

  static void requireWorkbookProtectionShape(WorkbookProtectionReport protection) {
    WorkbookInvariantChecks.require(protection != null, "workbook protection must not be null");
  }

  static void requirePackageSecurityShape(OoxmlPackageSecurityReport security) {
    WorkbookInvariantChecks.require(security != null, "package security must not be null");
    WorkbookInvariantChecks.require(
        security.encryption() != null, "package encryption must not be null");
    if (security.encryption().encrypted()) {
      WorkbookInvariantChecks.requirePresent(
          security.encryption().mode(), "encrypted package mode");
      WorkbookInvariantChecks.requirePresent(
          security.encryption().cipherAlgorithm(), "encrypted package cipherAlgorithm");
      WorkbookInvariantChecks.requirePresent(
          security.encryption().hashAlgorithm(), "encrypted package hashAlgorithm");
      WorkbookInvariantChecks.requirePresent(
          security.encryption().chainingMode(), "encrypted package chainingMode");
      WorkbookInvariantChecks.require(
          WorkbookInvariantChecks.requirePresent(
                  security.encryption().keyBits(), "encrypted package keyBits")
              > 0,
          "encrypted package keyBits must be positive");
      WorkbookInvariantChecks.require(
          WorkbookInvariantChecks.requirePresent(
                  security.encryption().blockSize(), "encrypted package blockSize")
              > 0,
          "encrypted package blockSize must be positive");
      WorkbookInvariantChecks.require(
          WorkbookInvariantChecks.requirePresent(
                  security.encryption().spinCount(), "encrypted package spinCount")
              >= 0,
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
        .forEach(WorkbookInvariantCellMetadataChecks::requireCustomXmlLinkedCellShape);
    mapping
        .linkedTables()
        .forEach(WorkbookInvariantCellMetadataChecks::requireCustomXmlLinkedTableShape);
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
          WorkbookInvariantCellStyleChecks.requireCellColorShape(color.color(), "autofilter color");
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
        .forEach(WorkbookInvariantCellMetadataChecks::requireAutofilterSortConditionShape);
  }

  static void requireAutofilterSortConditionShape(AutofilterSortConditionReport condition) {
    WorkbookInvariantChecks.require(
        condition != null, "autofilter sort condition must not be null");
    WorkbookInvariantChecks.requireNonBlank(condition.range(), "autofilter sort condition range");
    switch (condition) {
      case AutofilterSortConditionReport.Value _ -> {
        // No auxiliary payload.
      }
      case AutofilterSortConditionReport.CellColor color -> {
        WorkbookInvariantCellStyleChecks.requireCellColorShape(
            color.color(), "autofilter sort color");
      }
      case AutofilterSortConditionReport.FontColor color -> {
        WorkbookInvariantCellStyleChecks.requireCellColorShape(
            color.color(), "autofilter sort color");
      }
      case AutofilterSortConditionReport.Icon icon -> {
        WorkbookInvariantChecks.require(
            icon.iconId() >= 0, "autofilter sort iconId must not be negative");
      }
    }
  }

  static void requireTableColumnShape(TableColumnReport column) {
    WorkbookInvariantChecks.require(column != null, "table column must not be null");
    WorkbookInvariantChecks.require(column.id() >= 0L, "table column id must not be negative");
    WorkbookInvariantChecks.require(column.name() != null, "table column name must not be null");
  }
}
