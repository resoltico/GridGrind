package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PaneInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PrintAreaInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.PrintScalingInput;
import dev.erst.gridgrind.contract.dto.PrintSetupInput;
import dev.erst.gridgrind.contract.dto.PrintTitleColumnsInput;
import dev.erst.gridgrind.contract.dto.PrintTitleRowsInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterColumn;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterCriterion;
import dev.erst.gridgrind.excel.ExcelAutofilterSortCondition;
import dev.erst.gridgrind.excel.ExcelAutofilterSortState;
import dev.erst.gridgrind.excel.ExcelBinaryData;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold;
import dev.erst.gridgrind.excel.ExcelCustomXmlImportDefinition;
import dev.erst.gridgrind.excel.ExcelCustomXmlMappingLocator;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialBorderSide;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.ExcelHeaderFooterText;
import dev.erst.gridgrind.excel.ExcelIgnoredError;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.ExcelPivotTableDefinition;
import dev.erst.gridgrind.excel.ExcelPrintLayout;
import dev.erst.gridgrind.excel.ExcelPrintMargins;
import dev.erst.gridgrind.excel.ExcelPrintSetup;
import dev.erst.gridgrind.excel.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetDefaults;
import dev.erst.gridgrind.excel.ExcelSheetDisplay;
import dev.erst.gridgrind.excel.ExcelSheetOutlineSummary;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelSignatureLineDefinition;
import dev.erst.gridgrind.excel.ExcelTableColumnDefinition;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;
import java.util.Optional;

/**
 * Converts structured contract inputs such as drawings, validations, tables, and layout settings
 * into workbook-core records.
 *
 * <p>This helper intentionally spans the structured mutation surface, so PMD's import-count
 * heuristic is not a useful coupling signal here.
 */
@SuppressWarnings("PMD.ExcessiveImports")
final class WorkbookCommandStructuredInputConverter {
  private WorkbookCommandStructuredInputConverter() {}

  static ExcelCustomXmlImportDefinition toExcelCustomXmlImportDefinition(
      CustomXmlImportInput input) {
    return new ExcelCustomXmlImportDefinition(
        toExcelCustomXmlMappingLocator(input.locator()),
        WorkbookCommandSourceSupport.inlineText(input.xml(), "custom XML"));
  }

  static ExcelCustomXmlMappingLocator toExcelCustomXmlMappingLocator(
      dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator locator) {
    return new ExcelCustomXmlMappingLocator(locator.mapId(), locator.name());
  }

  static ExcelPictureDefinition toExcelPictureDefinition(PictureInput picture) {
    return new ExcelPictureDefinition(
        picture.name(),
        new ExcelBinaryData(
            WorkbookCommandSourceSupport.inlineBinary(picture.image().source(), "picture payload")),
        picture.image().format(),
        toExcelDrawingAnchor(picture.anchor()),
        picture.description() == null
            ? null
            : WorkbookCommandSourceSupport.inlineText(
                picture.description(), "picture description"));
  }

  static ExcelChartDefinition toExcelChartDefinition(ChartInput chart) {
    return new ExcelChartDefinition(
        chart.name(),
        toExcelDrawingAnchor(chart.anchor()),
        toExcelChartTitle(chart.title()),
        toExcelChartLegend(chart.legend()),
        chart.displayBlanksAs(),
        chart.plotOnlyVisibleCells(),
        chart.plots().stream()
            .map(WorkbookCommandStructuredInputConverter::toExcelChartPlot)
            .toList());
  }

  static ExcelSignatureLineDefinition toExcelSignatureLineDefinition(
      SignatureLineInput signatureLine) {
    return new ExcelSignatureLineDefinition(
        signatureLine.name(),
        toExcelDrawingAnchor(signatureLine.anchor()),
        signatureLine.allowComments(),
        signatureLine.signingInstructions(),
        signatureLine.suggestedSigner(),
        signatureLine.suggestedSigner2(),
        signatureLine.suggestedSignerEmail(),
        signatureLine.caption(),
        signatureLine.invalidStamp(),
        signatureLine.plainSignature() == null ? null : signatureLine.plainSignature().format(),
        signatureLine.plainSignature() == null
            ? null
            : new ExcelBinaryData(
                WorkbookCommandSourceSupport.inlineBinary(
                    signatureLine.plainSignature().source(), "signature-line plain signature")));
  }

  static ExcelShapeDefinition toExcelShapeDefinition(ShapeInput shape) {
    return new ExcelShapeDefinition(
        shape.name(),
        shape.kind(),
        toExcelDrawingAnchor(shape.anchor()),
        shape.presetGeometryToken(),
        shape.text() == null
            ? null
            : WorkbookCommandSourceSupport.inlineText(shape.text(), "shape text"));
  }

  static ExcelEmbeddedObjectDefinition toExcelEmbeddedObjectDefinition(
      EmbeddedObjectInput embeddedObject) {
    return new ExcelEmbeddedObjectDefinition(
        embeddedObject.name(),
        embeddedObject.label(),
        embeddedObject.fileName(),
        embeddedObject.command(),
        new ExcelBinaryData(
            WorkbookCommandSourceSupport.inlineBinary(
                embeddedObject.payload(), "embedded-object payload")),
        embeddedObject.previewImage().format(),
        new ExcelBinaryData(
            WorkbookCommandSourceSupport.inlineBinary(
                embeddedObject.previewImage().source(), "embedded-object preview image")),
        toExcelDrawingAnchor(embeddedObject.anchor()));
  }

  static ExcelDrawingAnchor.TwoCell toExcelDrawingAnchor(DrawingAnchorInput anchor) {
    return switch (anchor) {
      case DrawingAnchorInput.TwoCell twoCell ->
          new ExcelDrawingAnchor.TwoCell(
              new ExcelDrawingMarker(
                  twoCell.from().columnIndex(),
                  twoCell.from().rowIndex(),
                  twoCell.from().dx(),
                  twoCell.from().dy()),
              new ExcelDrawingMarker(
                  twoCell.to().columnIndex(),
                  twoCell.to().rowIndex(),
                  twoCell.to().dx(),
                  twoCell.to().dy()),
              twoCell.behavior());
    };
  }

  static ExcelDataValidationDefinition toExcelDataValidationDefinition(
      DataValidationInput validation) {
    return new ExcelDataValidationDefinition(
        toExcelDataValidationRule(validation.rule()),
        validation.allowBlank(),
        validation.suppressDropDownArrow(),
        toExcelDataValidationPrompt(validation.prompt()).orElse(null),
        toExcelDataValidationErrorAlert(validation.errorAlert()).orElse(null));
  }

  static ExcelDataValidationRule toExcelDataValidationRule(DataValidationRuleInput rule) {
    return switch (rule) {
      case DataValidationRuleInput.ExplicitList explicitList ->
          new ExcelDataValidationRule.ExplicitList(explicitList.values());
      case DataValidationRuleInput.FormulaList formulaList ->
          new ExcelDataValidationRule.FormulaList(formulaList.formula());
      case DataValidationRuleInput.WholeNumber wholeNumber ->
          new ExcelDataValidationRule.WholeNumber(
              wholeNumber.operator(), wholeNumber.formula1(), wholeNumber.formula2());
      case DataValidationRuleInput.DecimalNumber decimalNumber ->
          new ExcelDataValidationRule.DecimalNumber(
              decimalNumber.operator(), decimalNumber.formula1(), decimalNumber.formula2());
      case DataValidationRuleInput.DateRule dateRule ->
          new ExcelDataValidationRule.DateRule(
              dateRule.operator(), dateRule.formula1(), dateRule.formula2());
      case DataValidationRuleInput.TimeRule timeRule ->
          new ExcelDataValidationRule.TimeRule(
              timeRule.operator(), timeRule.formula1(), timeRule.formula2());
      case DataValidationRuleInput.TextLength textLength ->
          new ExcelDataValidationRule.TextLength(
              textLength.operator(), textLength.formula1(), textLength.formula2());
      case DataValidationRuleInput.CustomFormula customFormula ->
          new ExcelDataValidationRule.CustomFormula(customFormula.formula());
    };
  }

  static Optional<ExcelDataValidationPrompt> toExcelDataValidationPrompt(
      DataValidationPromptInput prompt) {
    return prompt == null
        ? Optional.empty()
        : Optional.of(
            new ExcelDataValidationPrompt(
                WorkbookCommandSourceSupport.inlineText(prompt.title(), "validation prompt title"),
                WorkbookCommandSourceSupport.inlineText(prompt.text(), "validation prompt text"),
                prompt.showPromptBox()));
  }

  static Optional<ExcelDataValidationErrorAlert> toExcelDataValidationErrorAlert(
      DataValidationErrorAlertInput errorAlert) {
    return errorAlert == null
        ? Optional.empty()
        : Optional.of(
            new ExcelDataValidationErrorAlert(
                errorAlert.style(),
                WorkbookCommandSourceSupport.inlineText(
                    errorAlert.title(), "validation error title"),
                WorkbookCommandSourceSupport.inlineText(errorAlert.text(), "validation error text"),
                errorAlert.showErrorBox()));
  }

  static ExcelConditionalFormattingBlockDefinition toExcelConditionalFormattingBlock(
      ConditionalFormattingBlockInput block) {
    return new ExcelConditionalFormattingBlockDefinition(
        block.ranges(),
        block.rules().stream()
            .map(WorkbookCommandStructuredInputConverter::toExcelConditionalFormattingRule)
            .toList());
  }

  static ExcelConditionalFormattingRule toExcelConditionalFormattingRule(
      ConditionalFormattingRuleInput rule) {
    return switch (rule) {
      case ConditionalFormattingRuleInput.FormulaRule formulaRule ->
          new ExcelConditionalFormattingRule.FormulaRule(
              formulaRule.formula(),
              formulaRule.stopIfTrue(),
              toExcelDifferentialStyle(formulaRule.style()).orElse(null));
      case ConditionalFormattingRuleInput.CellValueRule cellValueRule ->
          new ExcelConditionalFormattingRule.CellValueRule(
              cellValueRule.operator(),
              cellValueRule.formula1(),
              cellValueRule.formula2(),
              cellValueRule.stopIfTrue(),
              toExcelDifferentialStyle(cellValueRule.style()).orElse(null));
      case ConditionalFormattingRuleInput.ColorScaleRule colorScaleRule ->
          new ExcelConditionalFormattingRule.ColorScaleRule(
              colorScaleRule.thresholds().stream()
                  .map(
                      WorkbookCommandStructuredInputConverter
                          ::toExcelConditionalFormattingThreshold)
                  .toList(),
              colorScaleRule.colors().stream()
                  .map(
                      color ->
                          WorkbookCommandCellInputConverter.toRequiredExcelColor(
                              color, "color-scale color"))
                  .toList(),
              colorScaleRule.stopIfTrue());
      case ConditionalFormattingRuleInput.DataBarRule dataBarRule ->
          new ExcelConditionalFormattingRule.DataBarRule(
              WorkbookCommandCellInputConverter.toRequiredExcelColor(
                  dataBarRule.color(), "data-bar color"),
              dataBarRule.iconOnly(),
              dataBarRule.widthMin(),
              dataBarRule.widthMax(),
              toExcelConditionalFormattingThreshold(dataBarRule.minThreshold()),
              toExcelConditionalFormattingThreshold(dataBarRule.maxThreshold()),
              dataBarRule.stopIfTrue());
      case ConditionalFormattingRuleInput.IconSetRule iconSetRule ->
          new ExcelConditionalFormattingRule.IconSetRule(
              iconSetRule.iconSet(),
              iconSetRule.iconOnly(),
              iconSetRule.reversed(),
              iconSetRule.thresholds().stream()
                  .map(
                      WorkbookCommandStructuredInputConverter
                          ::toExcelConditionalFormattingThreshold)
                  .toList(),
              iconSetRule.stopIfTrue());
      case ConditionalFormattingRuleInput.Top10Rule top10Rule ->
          new ExcelConditionalFormattingRule.Top10Rule(
              top10Rule.rank(),
              top10Rule.percent(),
              top10Rule.bottom(),
              top10Rule.stopIfTrue(),
              toExcelDifferentialStyle(top10Rule.style()).orElse(null));
    };
  }

  static Optional<ExcelDifferentialStyle> toExcelDifferentialStyle(DifferentialStyleInput style) {
    if (style == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelDifferentialStyle(
            style.numberFormat(),
            style.bold(),
            style.italic(),
            WorkbookCommandCellInputConverter.toExcelFontHeight(style.fontHeight()).orElse(null),
            style.fontColor(),
            style.underline(),
            style.strikeout(),
            style.fillColor(),
            toExcelDifferentialBorder(style.border()).orElse(null)));
  }

  static Optional<ExcelDifferentialBorder> toExcelDifferentialBorder(
      DifferentialBorderInput border) {
    if (border == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelDifferentialBorder(
            toExcelDifferentialBorderSide(border.all()).orElse(null),
            toExcelDifferentialBorderSide(border.top()).orElse(null),
            toExcelDifferentialBorderSide(border.right()).orElse(null),
            toExcelDifferentialBorderSide(border.bottom()).orElse(null),
            toExcelDifferentialBorderSide(border.left()).orElse(null)));
  }

  static ExcelTableDefinition toExcelTableDefinition(TableInput table) {
    return new ExcelTableDefinition(
        table.name(),
        table.sheetName(),
        table.range(),
        table.showTotalsRow(),
        table.hasAutofilter(),
        toExcelTableStyle(table.style()),
        WorkbookCommandSourceSupport.inlineText(table.comment(), "table comment"),
        table.published(),
        table.insertRow(),
        table.insertRowShift(),
        table.headerRowCellStyle(),
        table.dataCellStyle(),
        table.totalsRowCellStyle(),
        table.columns().stream()
            .map(
                column ->
                    new ExcelTableColumnDefinition(
                        column.columnIndex(),
                        column.uniqueName(),
                        column.totalsRowLabel(),
                        column.totalsRowFunction(),
                        column.calculatedColumnFormula()))
            .toList());
  }

  static ExcelPivotTableDefinition toExcelPivotTableDefinition(PivotTableInput pivotTable) {
    return new ExcelPivotTableDefinition(
        pivotTable.name(),
        pivotTable.sheetName(),
        toExcelPivotTableSource(pivotTable.source()),
        new ExcelPivotTableDefinition.Anchor(pivotTable.anchor().topLeftAddress()),
        pivotTable.rowLabels(),
        pivotTable.columnLabels(),
        pivotTable.reportFilters(),
        pivotTable.dataFields().stream()
            .map(WorkbookCommandStructuredInputConverter::toExcelPivotTableDataField)
            .toList());
  }

  static ExcelTableStyle toExcelTableStyle(TableStyleInput style) {
    return switch (style) {
      case TableStyleInput.None _ -> new ExcelTableStyle.None();
      case TableStyleInput.Named named ->
          new ExcelTableStyle.Named(
              named.name(),
              named.showFirstColumn(),
              named.showLastColumn(),
              named.showRowStripes(),
              named.showColumnStripes());
    };
  }

  static ExcelNamedRangeScope toExcelNamedRangeScope(NamedRangeScope scope) {
    return switch (scope) {
      case NamedRangeScope.Workbook _ -> new ExcelNamedRangeScope.WorkbookScope();
      case NamedRangeScope.Sheet sheet -> new ExcelNamedRangeScope.SheetScope(sheet.sheetName());
    };
  }

  static ExcelNamedRangeScope toExcelNamedRangeScope(NamedRangeSelector.ScopedExact selector) {
    return switch (selector) {
      case NamedRangeSelector.WorkbookScope _ -> new ExcelNamedRangeScope.WorkbookScope();
      case NamedRangeSelector.SheetScope sheet ->
          new ExcelNamedRangeScope.SheetScope(sheet.sheetName());
    };
  }

  static String toExcelNamedRangeName(NamedRangeSelector.ScopedExact selector) {
    return switch (selector) {
      case NamedRangeSelector.WorkbookScope workbook -> workbook.name();
      case NamedRangeSelector.SheetScope sheet -> sheet.name();
    };
  }

  static ExcelNamedRangeTarget toExcelNamedRangeTarget(NamedRangeTarget target) {
    return target.formula() != null
        ? new ExcelNamedRangeTarget(target.formula())
        : new ExcelNamedRangeTarget(target.sheetName(), target.range());
  }

  static ExcelSheetCopyPosition toExcelSheetCopyPosition(SheetCopyPosition position) {
    return switch (position) {
      case SheetCopyPosition.AppendAtEnd _ -> new ExcelSheetCopyPosition.AppendAtEnd();
      case SheetCopyPosition.AtIndex atIndex ->
          new ExcelSheetCopyPosition.AtIndex(atIndex.targetIndex());
    };
  }

  static ExcelSheetProtectionSettings toExcelSheetProtectionSettings(
      SheetProtectionSettings settings) {
    return new ExcelSheetProtectionSettings(
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

  static ExcelWorkbookProtectionSettings toExcelWorkbookProtectionSettings(
      WorkbookProtectionInput protection) {
    return new ExcelWorkbookProtectionSettings(
        protection.structureLocked(),
        protection.windowsLocked(),
        protection.revisionsLocked(),
        protection.workbookPassword(),
        protection.revisionsPassword());
  }

  static ExcelSheetPane toExcelSheetPane(PaneInput pane) {
    return switch (pane) {
      case PaneInput.None _ -> new ExcelSheetPane.None();
      case PaneInput.Frozen frozen ->
          new ExcelSheetPane.Frozen(
              frozen.splitColumn(), frozen.splitRow(), frozen.leftmostColumn(), frozen.topRow());
      case PaneInput.Split split ->
          new ExcelSheetPane.Split(
              split.xSplitPosition(),
              split.ySplitPosition(),
              split.leftmostColumn(),
              split.topRow(),
              split.activePane());
    };
  }

  static ExcelPrintLayout toExcelPrintLayout(PrintLayoutInput printLayout) {
    return new ExcelPrintLayout(
        toExcelPrintArea(printLayout.printArea()),
        printLayout.orientation(),
        toExcelPrintScaling(printLayout.scaling()),
        toExcelPrintTitleRows(printLayout.repeatingRows()),
        toExcelPrintTitleColumns(printLayout.repeatingColumns()),
        new ExcelHeaderFooterText(
            WorkbookCommandSourceSupport.inlineText(printLayout.header().left(), "header left"),
            WorkbookCommandSourceSupport.inlineText(printLayout.header().center(), "header center"),
            WorkbookCommandSourceSupport.inlineText(printLayout.header().right(), "header right")),
        new ExcelHeaderFooterText(
            WorkbookCommandSourceSupport.inlineText(printLayout.footer().left(), "footer left"),
            WorkbookCommandSourceSupport.inlineText(printLayout.footer().center(), "footer center"),
            WorkbookCommandSourceSupport.inlineText(printLayout.footer().right(), "footer right")),
        toExcelPrintSetup(printLayout.setup()));
  }

  static ExcelSheetPresentation toExcelSheetPresentation(SheetPresentationInput presentation) {
    return new ExcelSheetPresentation(
        new ExcelSheetDisplay(
            presentation.display().displayGridlines(),
            presentation.display().displayZeros(),
            presentation.display().displayRowColHeadings(),
            presentation.display().displayFormulas(),
            presentation.display().rightToLeft()),
        WorkbookCommandCellInputConverter.toExcelColor(presentation.tabColor()).orElse(null),
        new ExcelSheetOutlineSummary(
            presentation.outlineSummary().rowSumsBelow(),
            presentation.outlineSummary().rowSumsRight()),
        new ExcelSheetDefaults(
            presentation.sheetDefaults().defaultColumnWidth(),
            presentation.sheetDefaults().defaultRowHeightPoints()),
        presentation.ignoredErrors().stream()
            .map(
                ignoredError ->
                    new ExcelIgnoredError(ignoredError.range(), ignoredError.errorTypes()))
            .toList());
  }

  static ExcelAutofilterFilterColumn toExcelAutofilterFilterColumn(
      AutofilterFilterColumnInput column) {
    return new ExcelAutofilterFilterColumn(
        column.columnId(),
        column.showButton(),
        toExcelAutofilterFilterCriterion(column.criterion()));
  }

  static Optional<ExcelAutofilterSortState> toExcelAutofilterSortState(
      AutofilterSortStateInput sortState) {
    if (sortState == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelAutofilterSortState(
            sortState.range(),
            sortState.caseSensitive(),
            sortState.columnSort(),
            sortState.sortMethod(),
            sortState.conditions().stream()
                .map(WorkbookCommandStructuredInputConverter::toExcelAutofilterSortCondition)
                .toList()));
  }

  private static ExcelChartDefinition.Title toExcelChartTitle(ChartInput.Title title) {
    return switch (title) {
      case ChartInput.Title.None _ -> new ExcelChartDefinition.Title.None();
      case ChartInput.Title.Text text ->
          new ExcelChartDefinition.Title.Text(
              WorkbookCommandSourceSupport.inlineText(text.source(), "chart title"));
      case ChartInput.Title.Formula formula ->
          new ExcelChartDefinition.Title.Formula(formula.formula());
    };
  }

  private static ExcelChartDefinition.Legend toExcelChartLegend(ChartInput.Legend legend) {
    return switch (legend) {
      case ChartInput.Legend.Hidden _ -> new ExcelChartDefinition.Legend.Hidden();
      case ChartInput.Legend.Visible visible ->
          new ExcelChartDefinition.Legend.Visible(visible.position());
    };
  }

  private static ExcelChartDefinition.Series toExcelChartSeries(ChartInput.Series series) {
    return new ExcelChartDefinition.Series(
        toExcelChartTitle(series.title()),
        toExcelChartDataSource(series.categories()),
        toExcelChartDataSource(series.values()),
        series.smooth(),
        series.markerStyle(),
        series.markerSize(),
        series.explosion());
  }

  private static ExcelChartDefinition.DataSource toExcelChartDataSource(
      ChartInput.DataSource source) {
    return switch (source) {
      case ChartInput.DataSource.Reference reference ->
          new ExcelChartDefinition.DataSource.Reference(reference.formula());
      case ChartInput.DataSource.StringLiteral literal ->
          new ExcelChartDefinition.DataSource.StringLiteral(literal.values());
      case ChartInput.DataSource.NumericLiteral literal ->
          new ExcelChartDefinition.DataSource.NumericLiteral(literal.values());
    };
  }

  private static ExcelChartDefinition.Axis toExcelChartAxis(ChartInput.Axis axis) {
    return new ExcelChartDefinition.Axis(
        axis.kind(), axis.position(), axis.crosses(), axis.visible());
  }

  private static ExcelChartDefinition.Plot toExcelChartPlot(ChartInput.Plot plot) {
    return switch (plot) {
      case ChartInput.Area area ->
          new ExcelChartDefinition.Area(
              area.varyColors(),
              area.grouping(),
              area.axes().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartAxis)
                  .toList(),
              area.series().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartSeries)
                  .toList());
      case ChartInput.Area3D area3D ->
          new ExcelChartDefinition.Area3D(
              area3D.varyColors(),
              area3D.grouping(),
              area3D.gapDepth(),
              area3D.axes().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartAxis)
                  .toList(),
              area3D.series().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartSeries)
                  .toList());
      case ChartInput.Bar bar ->
          new ExcelChartDefinition.Bar(
              bar.varyColors(),
              bar.barDirection(),
              bar.grouping(),
              bar.gapWidth(),
              bar.overlap(),
              bar.axes().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartAxis)
                  .toList(),
              bar.series().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartSeries)
                  .toList());
      case ChartInput.Bar3D bar3D ->
          new ExcelChartDefinition.Bar3D(
              bar3D.varyColors(),
              bar3D.barDirection(),
              bar3D.grouping(),
              bar3D.gapDepth(),
              bar3D.gapWidth(),
              bar3D.shape(),
              bar3D.axes().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartAxis)
                  .toList(),
              bar3D.series().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartSeries)
                  .toList());
      case ChartInput.Doughnut doughnut ->
          new ExcelChartDefinition.Doughnut(
              doughnut.varyColors(),
              doughnut.firstSliceAngle(),
              doughnut.holeSize(),
              doughnut.series().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartSeries)
                  .toList());
      case ChartInput.Line line ->
          new ExcelChartDefinition.Line(
              line.varyColors(),
              line.grouping(),
              line.axes().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartAxis)
                  .toList(),
              line.series().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartSeries)
                  .toList());
      case ChartInput.Line3D line3D ->
          new ExcelChartDefinition.Line3D(
              line3D.varyColors(),
              line3D.grouping(),
              line3D.gapDepth(),
              line3D.axes().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartAxis)
                  .toList(),
              line3D.series().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartSeries)
                  .toList());
      case ChartInput.Pie pie ->
          new ExcelChartDefinition.Pie(
              pie.varyColors(),
              pie.firstSliceAngle(),
              pie.series().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartSeries)
                  .toList());
      case ChartInput.Pie3D pie3D ->
          new ExcelChartDefinition.Pie3D(
              pie3D.varyColors(),
              pie3D.series().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartSeries)
                  .toList());
      case ChartInput.Radar radar ->
          new ExcelChartDefinition.Radar(
              radar.varyColors(),
              radar.style(),
              radar.axes().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartAxis)
                  .toList(),
              radar.series().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartSeries)
                  .toList());
      case ChartInput.Scatter scatter ->
          new ExcelChartDefinition.Scatter(
              scatter.varyColors(),
              scatter.style(),
              scatter.axes().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartAxis)
                  .toList(),
              scatter.series().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartSeries)
                  .toList());
      case ChartInput.Surface surface ->
          new ExcelChartDefinition.Surface(
              surface.varyColors(),
              surface.wireframe(),
              surface.axes().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartAxis)
                  .toList(),
              surface.series().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartSeries)
                  .toList());
      case ChartInput.Surface3D surface3D ->
          new ExcelChartDefinition.Surface3D(
              surface3D.varyColors(),
              surface3D.wireframe(),
              surface3D.axes().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartAxis)
                  .toList(),
              surface3D.series().stream()
                  .map(WorkbookCommandStructuredInputConverter::toExcelChartSeries)
                  .toList());
    };
  }

  private static ExcelAutofilterFilterCriterion toExcelAutofilterFilterCriterion(
      AutofilterFilterCriterionInput criterion) {
    return switch (criterion) {
      case AutofilterFilterCriterionInput.Values values ->
          new ExcelAutofilterFilterCriterion.Values(values.values(), values.includeBlank());
      case AutofilterFilterCriterionInput.Custom custom ->
          new ExcelAutofilterFilterCriterion.Custom(
              custom.and(),
              custom.conditions().stream()
                  .map(
                      condition ->
                          new ExcelAutofilterFilterCriterion.CustomCondition(
                              condition.operator(), condition.value()))
                  .toList());
      case AutofilterFilterCriterionInput.Dynamic dynamic ->
          new ExcelAutofilterFilterCriterion.Dynamic(
              dynamic.type(), dynamic.value(), dynamic.maxValue());
      case AutofilterFilterCriterionInput.Top10 top10 ->
          new ExcelAutofilterFilterCriterion.Top10(top10.value(), top10.top(), top10.percent());
      case AutofilterFilterCriterionInput.Color color ->
          new ExcelAutofilterFilterCriterion.Color(
              color.cellColor(),
              WorkbookCommandCellInputConverter.toRequiredExcelColor(
                  color.color(), "autofilter color"));
      case AutofilterFilterCriterionInput.Icon icon ->
          new ExcelAutofilterFilterCriterion.Icon(icon.iconSet(), icon.iconId());
    };
  }

  private static ExcelAutofilterSortCondition toExcelAutofilterSortCondition(
      AutofilterSortConditionInput condition) {
    return new ExcelAutofilterSortCondition(
        condition.range(),
        condition.descending(),
        condition.sortBy(),
        WorkbookCommandCellInputConverter.toExcelColor(condition.color()).orElse(null),
        condition.iconId());
  }

  private static ExcelConditionalFormattingThreshold toExcelConditionalFormattingThreshold(
      ConditionalFormattingThresholdInput threshold) {
    return new ExcelConditionalFormattingThreshold(
        threshold.type(), threshold.formula(), threshold.value());
  }

  private static Optional<ExcelDifferentialBorderSide> toExcelDifferentialBorderSide(
      DifferentialBorderSideInput side) {
    return side == null
        ? Optional.empty()
        : Optional.of(new ExcelDifferentialBorderSide(side.style(), side.color()));
  }

  private static ExcelPivotTableDefinition.Source toExcelPivotTableSource(
      PivotTableInput.Source source) {
    return switch (source) {
      case PivotTableInput.Source.Range range ->
          new ExcelPivotTableDefinition.Source.Range(range.sheetName(), range.range());
      case PivotTableInput.Source.NamedRange namedRange ->
          new ExcelPivotTableDefinition.Source.NamedRange(namedRange.name());
      case PivotTableInput.Source.Table table ->
          new ExcelPivotTableDefinition.Source.Table(table.name());
    };
  }

  private static ExcelPivotTableDefinition.DataField toExcelPivotTableDataField(
      PivotTableInput.DataField dataField) {
    return new ExcelPivotTableDefinition.DataField(
        dataField.sourceColumnName(),
        dataField.function(),
        dataField.displayName(),
        dataField.valueFormat());
  }

  private static ExcelPrintSetup toExcelPrintSetup(PrintSetupInput setup) {
    return new ExcelPrintSetup(
        new ExcelPrintMargins(
            setup.margins().left(),
            setup.margins().right(),
            setup.margins().top(),
            setup.margins().bottom(),
            setup.margins().header(),
            setup.margins().footer()),
        setup.printGridlines(),
        setup.horizontallyCentered(),
        setup.verticallyCentered(),
        setup.paperSize(),
        setup.draft(),
        setup.blackAndWhite(),
        setup.copies(),
        setup.useFirstPageNumber(),
        setup.firstPageNumber(),
        setup.rowBreaks(),
        setup.columnBreaks());
  }

  private static ExcelPrintLayout.Area toExcelPrintArea(PrintAreaInput printArea) {
    return switch (printArea) {
      case PrintAreaInput.None _ -> new ExcelPrintLayout.Area.None();
      case PrintAreaInput.Range range -> new ExcelPrintLayout.Area.Range(range.range());
    };
  }

  private static ExcelPrintLayout.Scaling toExcelPrintScaling(PrintScalingInput scaling) {
    return switch (scaling) {
      case PrintScalingInput.Automatic _ -> new ExcelPrintLayout.Scaling.Automatic();
      case PrintScalingInput.Fit fit ->
          new ExcelPrintLayout.Scaling.Fit(fit.widthPages(), fit.heightPages());
    };
  }

  private static ExcelPrintLayout.TitleRows toExcelPrintTitleRows(
      PrintTitleRowsInput repeatingRows) {
    return switch (repeatingRows) {
      case PrintTitleRowsInput.None _ -> new ExcelPrintLayout.TitleRows.None();
      case PrintTitleRowsInput.Band band ->
          new ExcelPrintLayout.TitleRows.Band(band.firstRowIndex(), band.lastRowIndex());
    };
  }

  private static ExcelPrintLayout.TitleColumns toExcelPrintTitleColumns(
      PrintTitleColumnsInput repeatingColumns) {
    return switch (repeatingColumns) {
      case PrintTitleColumnsInput.None _ -> new ExcelPrintLayout.TitleColumns.None();
      case PrintTitleColumnsInput.Band band ->
          new ExcelPrintLayout.TitleColumns.Band(band.firstColumnIndex(), band.lastColumnIndex());
    };
  }
}
