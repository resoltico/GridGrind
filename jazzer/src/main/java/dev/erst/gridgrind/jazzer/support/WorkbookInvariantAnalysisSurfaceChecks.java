package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.AutofilterEntryReport;
import dev.erst.gridgrind.contract.dto.AutofilterHealthReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdReport;
import dev.erst.gridgrind.contract.dto.DifferentialBorderReport;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideReport;
import dev.erst.gridgrind.contract.dto.DifferentialStyleReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.PivotTableHealthReport;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.PrintLayoutReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.TableHealthReport;
import dev.erst.gridgrind.contract.dto.TableStyleReport;
import java.util.List;

/**
 * Owns invariant checks for analysis payloads, structured workbook reports, and layout surfaces.
 */
final class WorkbookInvariantAnalysisSurfaceChecks {
  private WorkbookInvariantAnalysisSurfaceChecks() {}

  static void requireWindowShape(GridGrindResponse.WindowReport window) {
    WorkbookInvariantChecks.require(
        window.sheetName() != null, "window sheetName must not be null");
    WorkbookInvariantChecks.require(
        !window.sheetName().isBlank(), "window sheetName must not be blank");
    WorkbookInvariantChecks.require(
        window.topLeftAddress() != null, "window topLeftAddress must not be null");
    WorkbookInvariantChecks.require(
        !window.topLeftAddress().isBlank(), "window topLeftAddress must not be blank");
    WorkbookInvariantChecks.require(window.rows() != null, "window rows must not be null");
    WorkbookInvariantChecks.require(
        window.rows().size() == window.rowCount(), "window rows size must match rowCount");
    window
        .rows()
        .forEach(
            row -> {
              WorkbookInvariantChecks.require(
                  row.rowIndex() >= 0, "window row index must not be negative");
              WorkbookInvariantChecks.require(
                  row.cells() != null, "window row cells must not be null");
              WorkbookInvariantChecks.require(
                  row.cells().size() == window.columnCount(),
                  "window row cells size must match columnCount");
              row.cells().forEach(WorkbookInvariantCellSurfaceChecks::requireCellReportShape);
            });
  }

  static void requireHyperlinkEntryShape(GridGrindResponse.CellHyperlinkReport hyperlink) {
    WorkbookInvariantChecks.require(
        hyperlink.address() != null, "hyperlink address must not be null");
    WorkbookInvariantChecks.require(
        !hyperlink.address().isBlank(), "hyperlink address must not be blank");
    WorkbookInvariantChecks.require(
        hyperlink.hyperlink() != null, "hyperlink metadata must not be null");
    WorkbookInvariantCellSurfaceChecks.requireHyperlinkShape(hyperlink.hyperlink());
  }

  static void requireCommentEntryShape(GridGrindResponse.CellCommentReport comment) {
    WorkbookInvariantChecks.require(comment.address() != null, "comment address must not be null");
    WorkbookInvariantChecks.require(
        !comment.address().isBlank(), "comment address must not be blank");
    WorkbookInvariantChecks.require(comment.comment() != null, "comment metadata must not be null");
    WorkbookInvariantCellSurfaceChecks.requireCommentReportShape(comment.comment());
  }

  static void requireSheetLayoutShape(GridGrindResponse.SheetLayoutReport layout) {
    WorkbookInvariantChecks.require(
        layout.sheetName() != null, "layout sheetName must not be null");
    WorkbookInvariantChecks.require(
        !layout.sheetName().isBlank(), "layout sheetName must not be blank");
    WorkbookInvariantChecks.require(layout.pane() != null, "pane must not be null");
    WorkbookInvariantChecks.require(
        layout.zoomPercent() >= 10 && layout.zoomPercent() <= 400,
        "zoomPercent must be between 10 and 400 inclusive");
    switch (layout.pane()) {
      case dev.erst.gridgrind.contract.dto.PaneReport.None _ -> {}
      case dev.erst.gridgrind.contract.dto.PaneReport.Frozen frozen -> {
        WorkbookInvariantChecks.require(
            frozen.splitColumn() >= 0, "splitColumn must not be negative");
        WorkbookInvariantChecks.require(frozen.splitRow() >= 0, "splitRow must not be negative");
        WorkbookInvariantChecks.require(
            frozen.leftmostColumn() >= 0, "leftmostColumn must not be negative");
        WorkbookInvariantChecks.require(frozen.topRow() >= 0, "topRow must not be negative");
      }
      case dev.erst.gridgrind.contract.dto.PaneReport.Split split -> {
        WorkbookInvariantChecks.require(
            split.xSplitPosition() >= 0, "xSplitPosition must not be negative");
        WorkbookInvariantChecks.require(
            split.ySplitPosition() >= 0, "ySplitPosition must not be negative");
        WorkbookInvariantChecks.require(
            split.leftmostColumn() >= 0, "leftmostColumn must not be negative");
        WorkbookInvariantChecks.require(split.topRow() >= 0, "topRow must not be negative");
        WorkbookInvariantChecks.require(split.activePane() != null, "activePane must not be null");
      }
    }
    layout
        .columns()
        .forEach(
            column -> {
              WorkbookInvariantChecks.require(
                  column.columnIndex() >= 0, "columnIndex must not be negative");
              WorkbookInvariantChecks.require(
                  Double.isFinite(column.widthCharacters()) && column.widthCharacters() > 0.0d,
                  "column width must be finite and greater than 0");
            });
    layout
        .rows()
        .forEach(
            row -> {
              WorkbookInvariantChecks.require(row.rowIndex() >= 0, "rowIndex must not be negative");
              WorkbookInvariantChecks.require(
                  Double.isFinite(row.heightPoints()) && row.heightPoints() > 0.0d,
                  "row height must be finite and greater than 0");
            });
  }

  static void requirePrintLayoutShape(PrintLayoutReport layout) {
    WorkbookInvariantChecks.require(
        layout.sheetName() != null, "print layout sheetName must not be null");
    WorkbookInvariantChecks.require(
        !layout.sheetName().isBlank(), "print layout sheetName must not be blank");
    WorkbookInvariantChecks.require(layout.printArea() != null, "printArea must not be null");
    WorkbookInvariantChecks.require(layout.orientation() != null, "orientation must not be null");
    WorkbookInvariantChecks.require(layout.scaling() != null, "scaling must not be null");
    WorkbookInvariantChecks.require(
        layout.repeatingRows() != null, "repeatingRows must not be null");
    WorkbookInvariantChecks.require(
        layout.repeatingColumns() != null, "repeatingColumns must not be null");
    WorkbookInvariantChecks.require(layout.header() != null, "header must not be null");
    WorkbookInvariantChecks.require(layout.footer() != null, "footer must not be null");
    WorkbookInvariantCellSurfaceChecks.requirePrintSetupShape(layout.setup());
  }

  static void requireDataValidationEntryShape(
      dev.erst.gridgrind.contract.dto.DataValidationEntryReport validation) {
    WorkbookInvariantChecks.require(
        validation.ranges() != null, "data validation ranges must not be null");
    WorkbookInvariantChecks.require(
        !validation.ranges().isEmpty(), "data validation ranges must not be empty");
    validation
        .ranges()
        .forEach(range -> WorkbookInvariantChecks.requireNonBlank(range, "data validation range"));

    switch (validation) {
      case dev.erst.gridgrind.contract.dto.DataValidationEntryReport.Supported supported ->
          requireSupportedDataValidationShape(supported.validation());
      case dev.erst.gridgrind.contract.dto.DataValidationEntryReport.Unsupported unsupported -> {
        WorkbookInvariantChecks.requireNonBlank(unsupported.kind(), "data validation kind");
        WorkbookInvariantChecks.requireNonBlank(unsupported.detail(), "data validation detail");
      }
    }
  }

  static void requireAutofilterEntryShape(AutofilterEntryReport autofilter) {
    WorkbookInvariantChecks.requireNonBlank(autofilter.range(), "autofilter range");
    WorkbookInvariantChecks.require(
        autofilter.filterColumns() != null, "autofilter filterColumns must not be null");
    autofilter
        .filterColumns()
        .forEach(WorkbookInvariantCellSurfaceChecks::requireAutofilterFilterColumnShape);
    if (autofilter.sortState() != null) {
      WorkbookInvariantCellSurfaceChecks.requireAutofilterSortStateShape(autofilter.sortState());
    }
    switch (autofilter) {
      case AutofilterEntryReport.SheetOwned _ -> {}
      case AutofilterEntryReport.TableOwned tableOwned ->
          WorkbookInvariantChecks.requireNonBlank(tableOwned.tableName(), "autofilter table name");
    }
  }

  static void requireConditionalFormattingEntryShape(
      ConditionalFormattingEntryReport conditionalFormatting) {
    WorkbookInvariantChecks.require(
        conditionalFormatting.ranges() != null, "conditional formatting ranges must not be null");
    WorkbookInvariantChecks.require(
        !conditionalFormatting.ranges().isEmpty(),
        "conditional formatting ranges must not be empty");
    conditionalFormatting
        .ranges()
        .forEach(
            range ->
                WorkbookInvariantChecks.requireNonBlank(range, "conditional formatting range"));
    WorkbookInvariantChecks.require(
        conditionalFormatting.rules() != null, "conditional formatting rules must not be null");
    WorkbookInvariantChecks.require(
        !conditionalFormatting.rules().isEmpty(), "conditional formatting rules must not be empty");
    conditionalFormatting
        .rules()
        .forEach(WorkbookInvariantAnalysisSurfaceChecks::requireConditionalFormattingRuleShape);
  }

  static void requireTableEntryShape(TableEntryReport table) {
    WorkbookInvariantChecks.requireNonBlank(table.name(), "table name");
    WorkbookInvariantChecks.requireNonBlank(table.sheetName(), "table sheetName");
    WorkbookInvariantChecks.requireNonBlank(table.range(), "table range");
    WorkbookInvariantChecks.require(
        table.headerRowCount() >= 0, "table headerRowCount must not be negative");
    WorkbookInvariantChecks.require(
        table.totalsRowCount() >= 0, "table totalsRowCount must not be negative");
    WorkbookInvariantChecks.require(
        table.columnNames() != null, "table columnNames must not be null");
    WorkbookInvariantChecks.require(table.columns() != null, "table columns must not be null");
    WorkbookInvariantChecks.require(
        table.columnNames().size() == table.columns().size(),
        "table columnNames size must match columns size");
    table
        .columnNames()
        .forEach(
            columnName ->
                WorkbookInvariantChecks.require(
                    columnName != null, "table column name must not be null"));
    for (int index = 0; index < table.columns().size(); index++) {
      WorkbookInvariantCellSurfaceChecks.requireTableColumnShape(table.columns().get(index));
      WorkbookInvariantChecks.require(
          table.columnNames().get(index).equals(table.columns().get(index).name()),
          "table columnNames must align with columns");
    }
    WorkbookInvariantChecks.require(table.style() != null, "table style must not be null");
    requireTableStyleShape(table.style());
  }

  static void requirePivotTableShape(PivotTableReport pivotTable) {
    WorkbookInvariantWorkbookSurfaceChecks.requirePivotTableShape(pivotTable);
  }

  static void requireConditionalFormattingRuleShape(ConditionalFormattingRuleReport rule) {
    WorkbookInvariantChecks.require(
        rule.priority() > 0, "conditional formatting priority must be greater than 0");
    switch (rule) {
      case ConditionalFormattingRuleReport.FormulaRule formulaRule -> {
        WorkbookInvariantChecks.requireNonBlank(
            formulaRule.formula(), "conditional formatting formula");
        WorkbookInvariantChecks.require(
            formulaRule.style() != null, "conditional formatting style must not be null");
        requireDifferentialStyleShape(formulaRule.style());
      }
      case ConditionalFormattingRuleReport.CellValueRule cellValueRule -> {
        WorkbookInvariantChecks.require(
            cellValueRule.operator() != null, "conditional formatting operator must not be null");
        WorkbookInvariantChecks.requireNonBlank(
            cellValueRule.formula1(), "conditional formatting formula1");
        if (cellValueRule.formula2() != null) {
          WorkbookInvariantChecks.requireNonBlank(
              cellValueRule.formula2(), "conditional formatting formula2");
        }
        if (cellValueRule.style() != null) {
          requireDifferentialStyleShape(cellValueRule.style());
        }
      }
      case ConditionalFormattingRuleReport.ColorScaleRule colorScaleRule -> {
        WorkbookInvariantChecks.require(
            colorScaleRule.thresholds() != null,
            "conditional formatting thresholds must not be null");
        WorkbookInvariantChecks.require(
            !colorScaleRule.thresholds().isEmpty(),
            "conditional formatting thresholds must not be empty");
        colorScaleRule
            .thresholds()
            .forEach(
                WorkbookInvariantAnalysisSurfaceChecks::requireConditionalFormattingThresholdShape);
        WorkbookInvariantChecks.require(
            colorScaleRule.colors() != null, "conditional formatting colors must not be null");
        WorkbookInvariantChecks.require(
            !colorScaleRule.colors().isEmpty(), "conditional formatting colors must not be empty");
        colorScaleRule
            .colors()
            .forEach(
                color ->
                    WorkbookInvariantChecks.requireNonBlank(color, "conditional formatting color"));
      }
      case ConditionalFormattingRuleReport.DataBarRule dataBarRule -> {
        WorkbookInvariantChecks.requireNonBlank(
            dataBarRule.color(), "conditional formatting color");
        requireConditionalFormattingThresholdShape(dataBarRule.minThreshold());
        requireConditionalFormattingThresholdShape(dataBarRule.maxThreshold());
        WorkbookInvariantChecks.require(
            dataBarRule.widthMin() >= 0, "conditional formatting widthMin must not be negative");
        WorkbookInvariantChecks.require(
            dataBarRule.widthMax() >= 0, "conditional formatting widthMax must not be negative");
      }
      case ConditionalFormattingRuleReport.IconSetRule iconSetRule -> {
        WorkbookInvariantChecks.require(
            iconSetRule.iconSet() != null, "conditional formatting iconSet must not be null");
        WorkbookInvariantChecks.require(
            iconSetRule.thresholds() != null, "conditional formatting thresholds must not be null");
        WorkbookInvariantChecks.require(
            !iconSetRule.thresholds().isEmpty(),
            "conditional formatting thresholds must not be empty");
        iconSetRule
            .thresholds()
            .forEach(
                WorkbookInvariantAnalysisSurfaceChecks::requireConditionalFormattingThresholdShape);
      }
      case ConditionalFormattingRuleReport.Top10Rule top10Rule -> {
        WorkbookInvariantChecks.require(
            top10Rule.rank() >= 0, "conditional formatting rank must not be negative");
        if (top10Rule.style() != null) {
          requireDifferentialStyleShape(top10Rule.style());
        }
      }
      case ConditionalFormattingRuleReport.UnsupportedRule unsupportedRule -> {
        WorkbookInvariantChecks.requireNonBlank(
            unsupportedRule.kind(), "conditional formatting kind");
        WorkbookInvariantChecks.requireNonBlank(
            unsupportedRule.detail(), "conditional formatting detail");
      }
    }
  }

  static void requireConditionalFormattingThresholdShape(
      ConditionalFormattingThresholdReport threshold) {
    WorkbookInvariantChecks.require(
        threshold != null, "conditional formatting threshold must not be null");
    WorkbookInvariantChecks.require(
        threshold.type() != null, "conditional formatting threshold type must not be null");
  }

  static void requireDifferentialStyleShape(DifferentialStyleReport style) {
    WorkbookInvariantChecks.require(style != null, "conditional formatting style must not be null");
    if (style.numberFormat() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          style.numberFormat(), "conditional formatting numberFormat");
    }
    if (style.fontHeight() != null) {
      WorkbookInvariantCellSurfaceChecks.requireFontHeightShape(style.fontHeight());
    }
    style
        .fontColor()
        .ifPresent(
            color ->
                WorkbookInvariantChecks.requireNonBlank(color, "conditional formatting fontColor"));
    style
        .fillColor()
        .ifPresent(
            color ->
                WorkbookInvariantChecks.requireNonBlank(color, "conditional formatting fillColor"));
    style
        .border()
        .ifPresent(WorkbookInvariantAnalysisSurfaceChecks::requireDifferentialBorderShape);
    WorkbookInvariantChecks.require(
        style.unsupportedFeatures() != null,
        "conditional formatting unsupportedFeatures must not be null");
    style
        .unsupportedFeatures()
        .forEach(
            feature ->
                WorkbookInvariantChecks.require(
                    feature != null,
                    "conditional formatting unsupported feature must not be null"));
  }

  static void requireDifferentialBorderShape(DifferentialBorderReport border) {
    WorkbookInvariantChecks.require(
        border != null, "conditional formatting border must not be null");
    border
        .all()
        .ifPresent(WorkbookInvariantAnalysisSurfaceChecks::requireDifferentialBorderSideShape);
    border
        .top()
        .ifPresent(WorkbookInvariantAnalysisSurfaceChecks::requireDifferentialBorderSideShape);
    border
        .right()
        .ifPresent(WorkbookInvariantAnalysisSurfaceChecks::requireDifferentialBorderSideShape);
    border
        .bottom()
        .ifPresent(WorkbookInvariantAnalysisSurfaceChecks::requireDifferentialBorderSideShape);
    border
        .left()
        .ifPresent(WorkbookInvariantAnalysisSurfaceChecks::requireDifferentialBorderSideShape);
  }

  static void requireDifferentialBorderSideShape(DifferentialBorderSideReport side) {
    WorkbookInvariantChecks.require(
        side != null, "conditional formatting border side must not be null");
    WorkbookInvariantChecks.require(
        side.style() != null, "conditional formatting border style must not be null");
    side.color()
        .ifPresent(
            color ->
                WorkbookInvariantChecks.requireNonBlank(
                    color, "conditional formatting border color"));
  }

  static void requireTableStyleShape(TableStyleReport style) {
    switch (style) {
      case TableStyleReport.None _ -> {}
      case TableStyleReport.Named named ->
          WorkbookInvariantChecks.requireNonBlank(named.name(), "table style name");
    }
  }

  static void requireSupportedDataValidationShape(
      dev.erst.gridgrind.contract.dto.DataValidationEntryReport.DataValidationDefinitionReport
          validation) {
    WorkbookInvariantChecks.require(
        validation != null, "data validation definition must not be null");
    WorkbookInvariantChecks.require(
        validation.rule() != null, "data validation rule must not be null");
    switch (validation.rule()) {
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.ExplicitList explicitList -> {
        WorkbookInvariantChecks.require(
            explicitList.values() != null, "explicit list values must not be null");
        explicitList
            .values()
            .forEach(
                value -> WorkbookInvariantChecks.requireNonBlank(value, "explicit list value"));
      }
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.FormulaList formulaList ->
          WorkbookInvariantChecks.requireNonBlank(formulaList.formula(), "formula list formula");
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.WholeNumber wholeNumber ->
          requireComparisonRuleShape(wholeNumber.operator(), wholeNumber.formula1());
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.DecimalNumber decimalNumber ->
          requireComparisonRuleShape(decimalNumber.operator(), decimalNumber.formula1());
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.DateRule dateRule ->
          requireComparisonRuleShape(dateRule.operator(), dateRule.formula1());
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.TimeRule timeRule ->
          requireComparisonRuleShape(timeRule.operator(), timeRule.formula1());
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.TextLength textLength ->
          requireComparisonRuleShape(textLength.operator(), textLength.formula1());
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.CustomFormula customFormula ->
          WorkbookInvariantChecks.requireNonBlank(
              customFormula.formula(), "custom validation formula");
    }
    validation
        .prompt()
        .ifPresent(
            prompt -> {
              WorkbookInvariantChecks.requireNonBlank(
                  prompt.title(), "data validation prompt title");
              WorkbookInvariantChecks.requireNonBlank(prompt.text(), "data validation prompt text");
            });
    validation
        .errorAlert()
        .ifPresent(
            errorAlert -> {
              WorkbookInvariantChecks.require(
                  errorAlert.style() != null, "data validation error style must not be null");
              WorkbookInvariantChecks.requireNonBlank(
                  errorAlert.title(), "data validation error title");
              WorkbookInvariantChecks.requireNonBlank(
                  errorAlert.text(), "data validation error text");
            });
  }

  static void requireComparisonRuleShape(Object operator, String formula1) {
    WorkbookInvariantChecks.require(operator != null, "comparison operator must not be null");
    WorkbookInvariantChecks.requireNonBlank(formula1, "comparison formula1");
  }

  static void requireFormulaSurfaceShape(GridGrindResponse.FormulaSurfaceReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.totalFormulaCellCount() >= 0, "totalFormulaCellCount must not be negative");
    analysis
        .sheets()
        .forEach(
            sheet -> {
              WorkbookInvariantChecks.require(
                  sheet.sheetName() != null, "formula surface sheetName must not be null");
              WorkbookInvariantChecks.require(
                  !sheet.sheetName().isBlank(), "formula surface sheetName must not be blank");
              WorkbookInvariantChecks.require(
                  sheet.formulaCellCount() >= 0, "formulaCellCount must not be negative");
              WorkbookInvariantChecks.require(
                  sheet.distinctFormulaCount() >= 0, "distinctFormulaCount must not be negative");
              sheet
                  .formulas()
                  .forEach(
                      formula -> {
                        WorkbookInvariantChecks.require(
                            formula.formula() != null, "formula pattern must not be null");
                        WorkbookInvariantChecks.require(
                            !formula.formula().isBlank(), "formula pattern must not be blank");
                        WorkbookInvariantChecks.require(
                            formula.occurrenceCount() > 0,
                            "occurrenceCount must be greater than 0");
                        WorkbookInvariantChecks.require(
                            formula.addresses() != null, "formula addresses must not be null");
                      });
            });
  }

  static void requireSheetSchemaShape(GridGrindResponse.SheetSchemaReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.sheetName() != null, "schema sheetName must not be null");
    WorkbookInvariantChecks.require(
        !analysis.sheetName().isBlank(), "schema sheetName must not be blank");
    WorkbookInvariantChecks.require(
        analysis.topLeftAddress() != null, "schema topLeftAddress must not be null");
    WorkbookInvariantChecks.require(
        !analysis.topLeftAddress().isBlank(), "schema topLeftAddress must not be blank");
    WorkbookInvariantChecks.require(
        analysis.rowCount() > 0, "schema rowCount must be greater than 0");
    WorkbookInvariantChecks.require(
        analysis.columnCount() > 0, "schema columnCount must be greater than 0");
    WorkbookInvariantChecks.require(
        analysis.dataRowCount() >= 0, "schema dataRowCount must not be negative");
    analysis
        .columns()
        .forEach(
            column -> {
              WorkbookInvariantChecks.require(
                  column.columnIndex() >= 0, "schema columnIndex must not be negative");
              WorkbookInvariantChecks.require(
                  column.columnAddress() != null, "schema columnAddress must not be null");
              WorkbookInvariantChecks.require(
                  !column.columnAddress().isBlank(), "schema columnAddress must not be blank");
              WorkbookInvariantChecks.require(
                  column.headerDisplayValue() != null,
                  "schema headerDisplayValue must not be null");
              WorkbookInvariantChecks.require(
                  column.populatedCellCount() >= 0,
                  "schema populatedCellCount must not be negative");
              WorkbookInvariantChecks.require(
                  column.blankCellCount() >= 0, "schema blankCellCount must not be negative");
              column
                  .observedTypes()
                  .forEach(
                      typeCount -> {
                        WorkbookInvariantChecks.require(
                            typeCount.type() != null, "type count type must not be null");
                        WorkbookInvariantChecks.require(
                            !typeCount.type().isBlank(), "type count type must not be blank");
                        WorkbookInvariantChecks.require(
                            typeCount.count() > 0, "type count must be greater than 0");
                      });
            });
  }

  static void requireNamedRangeSurfaceShape(GridGrindResponse.NamedRangeSurfaceReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.workbookScopedCount() >= 0, "workbookScopedCount must not be negative");
    WorkbookInvariantChecks.require(
        analysis.sheetScopedCount() >= 0, "sheetScopedCount must not be negative");
    WorkbookInvariantChecks.require(
        analysis.rangeBackedCount() >= 0, "rangeBackedCount must not be negative");
    WorkbookInvariantChecks.require(
        analysis.formulaBackedCount() >= 0, "formulaBackedCount must not be negative");
    analysis
        .namedRanges()
        .forEach(
            namedRange -> {
              WorkbookInvariantChecks.require(
                  namedRange.name() != null, "named range name must not be null");
              WorkbookInvariantChecks.require(
                  !namedRange.name().isBlank(), "named range name must not be blank");
              WorkbookInvariantChecks.require(
                  namedRange.scope() != null, "named range scope must not be null");
              WorkbookInvariantChecks.require(
                  namedRange.refersToFormula() != null,
                  "named range refersToFormula must not be null");
              WorkbookInvariantChecks.require(
                  namedRange.kind() != null, "named range kind must not be null");
            });
  }

  static void requireFormulaHealthShape(GridGrindResponse.FormulaHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedFormulaCellCount() >= 0, "checkedFormulaCellCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireDataValidationHealthShape(
      dev.erst.gridgrind.contract.dto.DataValidationHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedValidationCount() >= 0, "checkedValidationCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireConditionalFormattingHealthShape(ConditionalFormattingHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedConditionalFormattingBlockCount() >= 0,
        "checkedConditionalFormattingBlockCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireAutofilterHealthShape(AutofilterHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedAutofilterCount() >= 0, "checkedAutofilterCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireTableHealthShape(TableHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedTableCount() >= 0, "checkedTableCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requirePivotTableHealthShape(PivotTableHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedPivotTableCount() >= 0, "checkedPivotTableCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireHyperlinkHealthShape(GridGrindResponse.HyperlinkHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedHyperlinkCount() >= 0, "checkedHyperlinkCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireNamedRangeHealthShape(GridGrindResponse.NamedRangeHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedNamedRangeCount() >= 0, "checkedNamedRangeCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireWorkbookFindingsShape(GridGrindResponse.WorkbookFindingsReport analysis) {
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireAnalysisSummaryShape(
      GridGrindResponse.AnalysisSummaryReport summary,
      List<GridGrindResponse.AnalysisFindingReport> findings) {
    WorkbookInvariantChecks.require(summary != null, "analysis summary must not be null");
    WorkbookInvariantChecks.require(findings != null, "analysis findings must not be null");
    WorkbookInvariantChecks.require(
        summary.totalCount() >= 0, "analysis totalCount must not be negative");
    WorkbookInvariantChecks.require(
        summary.errorCount() >= 0, "analysis errorCount must not be negative");
    WorkbookInvariantChecks.require(
        summary.warningCount() >= 0, "analysis warningCount must not be negative");
    WorkbookInvariantChecks.require(
        summary.infoCount() >= 0, "analysis infoCount must not be negative");
    WorkbookInvariantChecks.require(
        summary.totalCount() == findings.size(), "analysis totalCount must match findings size");
    WorkbookInvariantChecks.require(
        summary.totalCount() == summary.errorCount() + summary.warningCount() + summary.infoCount(),
        "analysis totalCount must equal error + warning + info");
    findings.forEach(WorkbookInvariantAnalysisSurfaceChecks::requireAnalysisFindingShape);
  }

  static void requireAnalysisFindingShape(GridGrindResponse.AnalysisFindingReport finding) {
    WorkbookInvariantChecks.require(
        finding.code() != null, "analysis finding code must not be null");
    WorkbookInvariantChecks.require(
        finding.severity() != null, "analysis finding severity must not be null");
    WorkbookInvariantChecks.requireNonBlank(finding.title(), "analysis title");
    WorkbookInvariantChecks.requireNonBlank(finding.message(), "analysis message");
    WorkbookInvariantChecks.require(
        finding.location() != null, "analysis location must not be null");
    WorkbookInvariantChecks.require(
        finding.evidence() != null, "analysis evidence must not be null");
    finding
        .evidence()
        .forEach(
            evidence -> WorkbookInvariantChecks.requireNonBlank(evidence, "analysis evidence"));

    switch (finding.location()) {
      case GridGrindResponse.AnalysisLocationReport.Workbook _ -> {}
      case GridGrindResponse.AnalysisLocationReport.Sheet sheet ->
          WorkbookInvariantChecks.requireNonBlank(sheet.sheetName(), "analysis sheetName");
      case GridGrindResponse.AnalysisLocationReport.Cell cell -> {
        WorkbookInvariantChecks.requireNonBlank(cell.sheetName(), "analysis sheetName");
        WorkbookInvariantChecks.requireNonBlank(cell.address(), "analysis address");
      }
      case GridGrindResponse.AnalysisLocationReport.Range range -> {
        WorkbookInvariantChecks.requireNonBlank(range.sheetName(), "analysis sheetName");
        WorkbookInvariantChecks.requireNonBlank(range.range(), "analysis range");
      }
      case GridGrindResponse.AnalysisLocationReport.NamedRange namedRange -> {
        WorkbookInvariantChecks.requireNonBlank(namedRange.name(), "analysis named range");
        WorkbookInvariantChecks.require(
            namedRange.scope() != null, "analysis named range scope must not be null");
      }
    }
  }
}
