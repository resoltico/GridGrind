package dev.erst.gridgrind.protocol.exec;

import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.util.List;

/** Converts workbook-core read results into protocol response shapes. */
final class WorkbookReadResultConverter {
  private WorkbookReadResultConverter() {}

  /** Converts one workbook-core read result into the protocol response shape. */
  static WorkbookReadResult toReadResult(dev.erst.gridgrind.excel.WorkbookReadResult result) {
    return switch (result) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult workbookSummary ->
          new WorkbookReadResult.WorkbookSummaryResult(
              workbookSummary.requestId(), toWorkbookSummary(workbookSummary.workbook()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookProtectionResult protection ->
          new WorkbookReadResult.WorkbookProtectionResult(
              protection.requestId(), toWorkbookProtectionReport(protection.protection()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangesResult namedRanges ->
          new WorkbookReadResult.NamedRangesResult(
              namedRanges.requestId(),
              namedRanges.namedRanges().stream()
                  .map(WorkbookReadResultConverter::toNamedRangeReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult sheetSummary ->
          new WorkbookReadResult.SheetSummaryResult(
              sheetSummary.requestId(),
              new GridGrindResponse.SheetSummaryReport(
                  sheetSummary.sheet().sheetName(),
                  sheetSummary.sheet().visibility(),
                  toSheetProtectionReport(sheetSummary.sheet().protection()),
                  sheetSummary.sheet().physicalRowCount(),
                  sheetSummary.sheet().lastRowIndex(),
                  sheetSummary.sheet().lastColumnIndex()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.CellsResult cells ->
          new WorkbookReadResult.CellsResult(
              cells.requestId(),
              cells.sheetName(),
              cells.cells().stream().map(WorkbookReadResultConverter::toCellReport).toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.WindowResult window ->
          new WorkbookReadResult.WindowResult(window.requestId(), toWindowReport(window.window()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.MergedRegionsResult mergedRegions ->
          new WorkbookReadResult.MergedRegionsResult(
              mergedRegions.requestId(),
              mergedRegions.sheetName(),
              mergedRegions.mergedRegions().stream()
                  .map(region -> new GridGrindResponse.MergedRegionReport(region.range()))
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinksResult hyperlinks ->
          new WorkbookReadResult.HyperlinksResult(
              hyperlinks.requestId(),
              hyperlinks.sheetName(),
              hyperlinks.hyperlinks().stream()
                  .map(WorkbookReadResultConverter::toCellHyperlinkReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.CommentsResult comments ->
          new WorkbookReadResult.CommentsResult(
              comments.requestId(),
              comments.sheetName(),
              comments.comments().stream()
                  .map(WorkbookReadResultConverter::toCellCommentReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult sheetLayout ->
          new WorkbookReadResult.SheetLayoutResult(
              sheetLayout.requestId(), toSheetLayoutReport(sheetLayout.layout()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult printLayout ->
          new WorkbookReadResult.PrintLayoutResult(
              printLayout.requestId(), toPrintLayoutReport(printLayout));
      case dev.erst.gridgrind.excel.WorkbookReadResult.DataValidationsResult dataValidations ->
          new WorkbookReadResult.DataValidationsResult(
              dataValidations.requestId(),
              dataValidations.sheetName(),
              dataValidations.validations().stream()
                  .map(WorkbookReadResultConverter::toDataValidationEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult
              conditionalFormatting ->
          new WorkbookReadResult.ConditionalFormattingResult(
              conditionalFormatting.requestId(),
              conditionalFormatting.sheetName(),
              conditionalFormatting.conditionalFormattingBlocks().stream()
                  .map(WorkbookReadResultConverter::toConditionalFormattingEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.AutofiltersResult autofilters ->
          new WorkbookReadResult.AutofiltersResult(
              autofilters.requestId(),
              autofilters.sheetName(),
              autofilters.autofilters().stream()
                  .map(WorkbookReadResultConverter::toAutofilterEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.TablesResult tables ->
          new WorkbookReadResult.TablesResult(
              tables.requestId(),
              tables.tables().stream()
                  .map(WorkbookReadResultConverter::toTableEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurfaceResult formulaSurface ->
          new WorkbookReadResult.FormulaSurfaceResult(
              formulaSurface.requestId(), toFormulaSurfaceReport(formulaSurface.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchemaResult sheetSchema ->
          new WorkbookReadResult.SheetSchemaResult(
              sheetSchema.requestId(), toSheetSchemaReport(sheetSchema.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceResult namedRangeSurface ->
          new WorkbookReadResult.NamedRangeSurfaceResult(
              namedRangeSurface.requestId(),
              toNamedRangeSurfaceReport(namedRangeSurface.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.FormulaHealthResult formulaHealth ->
          new WorkbookReadResult.FormulaHealthResult(
              formulaHealth.requestId(), toFormulaHealthReport(formulaHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.DataValidationHealthResult
              dataValidationHealth ->
          new WorkbookReadResult.DataValidationHealthResult(
              dataValidationHealth.requestId(),
              toDataValidationHealthReport(dataValidationHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingHealthResult
              conditionalFormattingHealth ->
          new WorkbookReadResult.ConditionalFormattingHealthResult(
              conditionalFormattingHealth.requestId(),
              toConditionalFormattingHealthReport(conditionalFormattingHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.AutofilterHealthResult autofilterHealth ->
          new WorkbookReadResult.AutofilterHealthResult(
              autofilterHealth.requestId(), toAutofilterHealthReport(autofilterHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.TableHealthResult tableHealth ->
          new WorkbookReadResult.TableHealthResult(
              tableHealth.requestId(), toTableHealthReport(tableHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinkHealthResult hyperlinkHealth ->
          new WorkbookReadResult.HyperlinkHealthResult(
              hyperlinkHealth.requestId(), toHyperlinkHealthReport(hyperlinkHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeHealthResult namedRangeHealth ->
          new WorkbookReadResult.NamedRangeHealthResult(
              namedRangeHealth.requestId(), toNamedRangeHealthReport(namedRangeHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookFindingsResult workbookFindings ->
          new WorkbookReadResult.WorkbookFindingsResult(
              workbookFindings.requestId(), toWorkbookFindingsReport(workbookFindings.analysis()));
    };
  }

  /** Converts one workbook-core workbook summary into the protocol response shape. */
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

  /** Converts one workbook-core named-range snapshot into the protocol response shape. */
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
    if (hyperlink == null) {
      return null;
    }
    return switch (hyperlink) {
      case ExcelHyperlink.Url url -> new HyperlinkTarget.Url(url.target());
      case ExcelHyperlink.Email email -> new HyperlinkTarget.Email(email.target());
      case ExcelHyperlink.File file -> new HyperlinkTarget.File(file.path());
      case ExcelHyperlink.Document document -> new HyperlinkTarget.Document(document.target());
    };
  }

  /** Converts workbook-core comment metadata into the protocol response shape. */
  static GridGrindResponse.CommentReport toCommentReport(ExcelComment comment) {
    if (comment == null) {
      return null;
    }
    return new GridGrindResponse.CommentReport(comment.text(), comment.author(), comment.visible());
  }

  /** Converts workbook-core workbook-protection state into the protocol response shape. */
  static WorkbookProtectionReport toWorkbookProtectionReport(
      ExcelWorkbookProtectionSnapshot protection) {
    return new WorkbookProtectionReport(
        protection.structureLocked(),
        protection.windowsLocked(),
        protection.revisionLocked(),
        protection.workbookPasswordHashPresent(),
        protection.revisionsPasswordHashPresent());
  }

  static DataValidationEntryReport toDataValidationEntryReport(
      ExcelDataValidationSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelDataValidationSnapshot.Supported supported ->
          new DataValidationEntryReport.Supported(
              supported.ranges(), toDataValidationDefinitionReport(supported.validation()));
      case ExcelDataValidationSnapshot.Unsupported unsupported ->
          new DataValidationEntryReport.Unsupported(
              unsupported.ranges(), unsupported.kind(), unsupported.detail());
    };
  }

  static DataValidationEntryReport.DataValidationDefinitionReport toDataValidationDefinitionReport(
      ExcelDataValidationDefinition definition) {
    return new DataValidationEntryReport.DataValidationDefinitionReport(
        toDataValidationRuleInput(definition.rule()),
        definition.allowBlank(),
        definition.suppressDropDownArrow(),
        definition.prompt() == null
            ? null
            : new DataValidationPromptInput(
                definition.prompt().title(),
                definition.prompt().text(),
                definition.prompt().showPromptBox()),
        definition.errorAlert() == null
            ? null
            : new DataValidationErrorAlertInput(
                definition.errorAlert().style(),
                definition.errorAlert().title(),
                definition.errorAlert().text(),
                definition.errorAlert().showErrorBox()));
  }

  private static DataValidationRuleInput toDataValidationRuleInput(ExcelDataValidationRule rule) {
    return switch (rule) {
      case ExcelDataValidationRule.ExplicitList explicitList ->
          new DataValidationRuleInput.ExplicitList(explicitList.values());
      case ExcelDataValidationRule.FormulaList formulaList ->
          new DataValidationRuleInput.FormulaList(formulaList.formula());
      case ExcelDataValidationRule.WholeNumber wholeNumber ->
          new DataValidationRuleInput.WholeNumber(
              wholeNumber.operator(), wholeNumber.formula1(), wholeNumber.formula2());
      case ExcelDataValidationRule.DecimalNumber decimalNumber ->
          new DataValidationRuleInput.DecimalNumber(
              decimalNumber.operator(), decimalNumber.formula1(), decimalNumber.formula2());
      case ExcelDataValidationRule.DateRule dateRule ->
          new DataValidationRuleInput.DateRule(
              dateRule.operator(), dateRule.formula1(), dateRule.formula2());
      case ExcelDataValidationRule.TimeRule timeRule ->
          new DataValidationRuleInput.TimeRule(
              timeRule.operator(), timeRule.formula1(), timeRule.formula2());
      case ExcelDataValidationRule.TextLength textLength ->
          new DataValidationRuleInput.TextLength(
              textLength.operator(), textLength.formula1(), textLength.formula2());
      case ExcelDataValidationRule.CustomFormula customFormula ->
          new DataValidationRuleInput.CustomFormula(customFormula.formula());
    };
  }

  static ConditionalFormattingEntryReport toConditionalFormattingEntryReport(
      ExcelConditionalFormattingBlockSnapshot block) {
    return new ConditionalFormattingEntryReport(
        block.ranges(),
        block.rules().stream()
            .map(WorkbookReadResultConverter::toConditionalFormattingRuleReport)
            .toList());
  }

  static ConditionalFormattingRuleReport toConditionalFormattingRuleReport(
      ExcelConditionalFormattingRuleSnapshot rule) {
    return switch (rule) {
      case ExcelConditionalFormattingRuleSnapshot.FormulaRule formulaRule ->
          new ConditionalFormattingRuleReport.FormulaRule(
              formulaRule.priority(),
              formulaRule.stopIfTrue(),
              formulaRule.formula(),
              toDifferentialStyleReport(formulaRule.style()));
      case ExcelConditionalFormattingRuleSnapshot.CellValueRule cellValueRule ->
          new ConditionalFormattingRuleReport.CellValueRule(
              cellValueRule.priority(),
              cellValueRule.stopIfTrue(),
              cellValueRule.operator(),
              cellValueRule.formula1(),
              cellValueRule.formula2(),
              toDifferentialStyleReport(cellValueRule.style()));
      case ExcelConditionalFormattingRuleSnapshot.ColorScaleRule colorScaleRule ->
          new ConditionalFormattingRuleReport.ColorScaleRule(
              colorScaleRule.priority(),
              colorScaleRule.stopIfTrue(),
              colorScaleRule.thresholds().stream()
                  .map(WorkbookReadResultConverter::toConditionalFormattingThresholdReport)
                  .toList(),
              colorScaleRule.colors());
      case ExcelConditionalFormattingRuleSnapshot.DataBarRule dataBarRule ->
          new ConditionalFormattingRuleReport.DataBarRule(
              dataBarRule.priority(),
              dataBarRule.stopIfTrue(),
              dataBarRule.color(),
              dataBarRule.iconOnly(),
              dataBarRule.widthMin(),
              dataBarRule.widthMax(),
              toConditionalFormattingThresholdReport(dataBarRule.minThreshold()),
              toConditionalFormattingThresholdReport(dataBarRule.maxThreshold()));
      case ExcelConditionalFormattingRuleSnapshot.IconSetRule iconSetRule ->
          new ConditionalFormattingRuleReport.IconSetRule(
              iconSetRule.priority(),
              iconSetRule.stopIfTrue(),
              iconSetRule.iconSet(),
              iconSetRule.iconOnly(),
              iconSetRule.reversed(),
              iconSetRule.thresholds().stream()
                  .map(WorkbookReadResultConverter::toConditionalFormattingThresholdReport)
                  .toList());
      case ExcelConditionalFormattingRuleSnapshot.Top10Rule top10Rule ->
          new ConditionalFormattingRuleReport.Top10Rule(
              top10Rule.priority(),
              top10Rule.stopIfTrue(),
              top10Rule.rank(),
              top10Rule.percent(),
              top10Rule.bottom(),
              toDifferentialStyleReport(top10Rule.style()));
      case ExcelConditionalFormattingRuleSnapshot.UnsupportedRule unsupportedRule ->
          new ConditionalFormattingRuleReport.UnsupportedRule(
              unsupportedRule.priority(),
              unsupportedRule.stopIfTrue(),
              unsupportedRule.kind(),
              unsupportedRule.detail());
    };
  }

  static ConditionalFormattingThresholdReport toConditionalFormattingThresholdReport(
      ExcelConditionalFormattingThresholdSnapshot threshold) {
    return new ConditionalFormattingThresholdReport(
        threshold.type(), threshold.formula(), threshold.value());
  }

  static DifferentialStyleReport toDifferentialStyleReport(ExcelDifferentialStyleSnapshot style) {
    if (style == null) {
      return null;
    }
    return new DifferentialStyleReport(
        style.numberFormat(),
        style.bold(),
        style.italic(),
        toFontHeightReport(style.fontHeight()),
        style.fontColor(),
        style.underline(),
        style.strikeout(),
        style.fillColor(),
        toDifferentialBorderReport(style.border()),
        style.unsupportedFeatures());
  }

  static DifferentialBorderReport toDifferentialBorderReport(ExcelDifferentialBorder border) {
    if (border == null) {
      return null;
    }
    return new DifferentialBorderReport(
        toDifferentialBorderSideReport(border.all()),
        toDifferentialBorderSideReport(border.top()),
        toDifferentialBorderSideReport(border.right()),
        toDifferentialBorderSideReport(border.bottom()),
        toDifferentialBorderSideReport(border.left()));
  }

  static DifferentialBorderSideReport toDifferentialBorderSideReport(
      ExcelDifferentialBorderSide side) {
    return side == null ? null : new DifferentialBorderSideReport(side.style(), side.color());
  }

  static FontHeightReport toFontHeightReport(ExcelFontHeight fontHeight) {
    return fontHeight == null
        ? null
        : new FontHeightReport(fontHeight.twips(), fontHeight.points());
  }

  static GridGrindResponse.CellStyleReport toCellStyleReport(ExcelCellStyleSnapshot style) {
    return new GridGrindResponse.CellStyleReport(
        style.numberFormat(),
        new CellAlignmentReport(
            style.alignment().wrapText(),
            style.alignment().horizontalAlignment(),
            style.alignment().verticalAlignment(),
            style.alignment().textRotation(),
            style.alignment().indentation()),
        toCellFontReport(style.font()),
        new CellFillReport(
            style.fill().pattern(),
            toCellColorReport(style.fill().foregroundColor()),
            toCellColorReport(style.fill().backgroundColor()),
            toCellGradientFillReport(style.fill().gradient())),
        new CellBorderReport(
            toCellBorderSideReport(style.border().top()),
            toCellBorderSideReport(style.border().right()),
            toCellBorderSideReport(style.border().bottom()),
            toCellBorderSideReport(style.border().left())),
        new CellProtectionReport(style.protection().locked(), style.protection().hiddenFormula()));
  }

  static CellFontReport toCellFontReport(ExcelCellFontSnapshot font) {
    return new CellFontReport(
        font.bold(),
        font.italic(),
        font.fontName(),
        toFontHeightReport(font.fontHeight()),
        toCellColorReport(font.fontColor()),
        font.underline(),
        font.strikeout());
  }

  static CellBorderSideReport toCellBorderSideReport(ExcelBorderSideSnapshot side) {
    return new CellBorderSideReport(side.style(), toCellColorReport(side.color()));
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
                        row.cells().stream()
                            .map(WorkbookReadResultConverter::toCellReport)
                            .toList()))
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

  private static GridGrindResponse.CellReport toCellReport(ExcelCellSnapshot snapshot) {
    GridGrindResponse.CellStyleReport style = toCellStyleReport(snapshot.style());
    HyperlinkTarget hyperlink = toHyperlinkTarget(snapshot.metadata().hyperlink().orElse(null));
    GridGrindResponse.CommentReport comment =
        toCommentReport(snapshot.metadata().comment().orElse(null));

    return switch (snapshot) {
      case ExcelCellSnapshot.BlankSnapshot s ->
          new GridGrindResponse.CellReport.BlankReport(
              s.address(), s.declaredType(), s.displayValue(), style, hyperlink, comment);
      case ExcelCellSnapshot.TextSnapshot s ->
          new GridGrindResponse.CellReport.TextReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.stringValue(),
              toRichTextRunReports(s.richText()));
      case ExcelCellSnapshot.NumberSnapshot s ->
          new GridGrindResponse.CellReport.NumberReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.numberValue());
      case ExcelCellSnapshot.BooleanSnapshot s ->
          new GridGrindResponse.CellReport.BooleanReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.booleanValue());
      case ExcelCellSnapshot.ErrorSnapshot s ->
          new GridGrindResponse.CellReport.ErrorReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.errorValue());
      case ExcelCellSnapshot.FormulaSnapshot s ->
          new GridGrindResponse.CellReport.FormulaReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.formula(),
              toCellReport(s.evaluation()));
    };
  }

  @SuppressWarnings("PMD.ReturnEmptyCollectionRatherThanNull")
  private static List<RichTextRunReport> toRichTextRunReports(ExcelRichTextSnapshot richText) {
    if (richText == null) {
      return null;
    }
    return richText.runs().stream()
        .map(run -> new RichTextRunReport(run.text(), toCellFontReport(run.font())))
        .toList();
  }

  private static CellColorReport toCellColorReport(ExcelColorSnapshot color) {
    return color == null
        ? null
        : new CellColorReport(color.rgb(), color.theme(), color.indexed(), color.tint());
  }

  private static CellGradientFillReport toCellGradientFillReport(
      ExcelGradientFillSnapshot gradient) {
    return gradient == null
        ? null
        : new CellGradientFillReport(
            gradient.type(),
            gradient.degree(),
            gradient.left(),
            gradient.right(),
            gradient.top(),
            gradient.bottom(),
            gradient.stops().stream()
                .map(
                    stop ->
                        new CellGradientStopReport(
                            stop.position(), toCellColorReport(stop.color())))
                .toList());
  }

  private static GridGrindResponse.FormulaSurfaceReport toFormulaSurfaceReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurface analysis) {
    return new GridGrindResponse.FormulaSurfaceReport(
        analysis.totalFormulaCellCount(),
        analysis.sheets().stream()
            .map(
                sheet ->
                    new GridGrindResponse.SheetFormulaSurfaceReport(
                        sheet.sheetName(),
                        sheet.formulaCellCount(),
                        sheet.distinctFormulaCount(),
                        sheet.formulas().stream()
                            .map(
                                formula ->
                                    new GridGrindResponse.FormulaPatternReport(
                                        formula.formula(),
                                        formula.occurrenceCount(),
                                        formula.addresses()))
                            .toList()))
            .toList());
  }

  private static GridGrindResponse.SheetSchemaReport toSheetSchemaReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchema analysis) {
    return new GridGrindResponse.SheetSchemaReport(
        analysis.sheetName(),
        analysis.topLeftAddress(),
        analysis.rowCount(),
        analysis.columnCount(),
        analysis.dataRowCount(),
        analysis.columns().stream()
            .map(
                column ->
                    new GridGrindResponse.SchemaColumnReport(
                        column.columnIndex(),
                        column.columnAddress(),
                        column.headerDisplayValue(),
                        column.populatedCellCount(),
                        column.blankCellCount(),
                        column.observedTypes().stream()
                            .map(
                                typeCount ->
                                    new GridGrindResponse.TypeCountReport(
                                        typeCount.type(), typeCount.count()))
                            .toList(),
                        column.dominantType()))
            .toList());
  }

  private static GridGrindResponse.NamedRangeSurfaceReport toNamedRangeSurfaceReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurface analysis) {
    return new GridGrindResponse.NamedRangeSurfaceReport(
        analysis.workbookScopedCount(),
        analysis.sheetScopedCount(),
        analysis.rangeBackedCount(),
        analysis.formulaBackedCount(),
        analysis.namedRanges().stream()
            .map(
                entry ->
                    new GridGrindResponse.NamedRangeSurfaceEntryReport(
                        entry.name(),
                        toNamedRangeScope(entry.scope()),
                        entry.refersToFormula(),
                        switch (entry.kind()) {
                          case RANGE -> GridGrindResponse.NamedRangeBackingKind.RANGE;
                          case FORMULA -> GridGrindResponse.NamedRangeBackingKind.FORMULA;
                        }))
            .toList());
  }

  private static GridGrindResponse.FormulaHealthReport toFormulaHealthReport(
      WorkbookAnalysis.FormulaHealth analysis) {
    return new GridGrindResponse.FormulaHealthReport(
        analysis.checkedFormulaCellCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(WorkbookReadResultConverter::toAnalysisFindingReport)
            .toList());
  }

  private static DataValidationHealthReport toDataValidationHealthReport(
      WorkbookAnalysis.DataValidationHealth analysis) {
    return new DataValidationHealthReport(
        analysis.checkedValidationCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(WorkbookReadResultConverter::toAnalysisFindingReport)
            .toList());
  }

  private static ConditionalFormattingHealthReport toConditionalFormattingHealthReport(
      WorkbookAnalysis.ConditionalFormattingHealth analysis) {
    return new ConditionalFormattingHealthReport(
        analysis.checkedConditionalFormattingBlockCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(WorkbookReadResultConverter::toAnalysisFindingReport)
            .toList());
  }

  private static AutofilterHealthReport toAutofilterHealthReport(
      WorkbookAnalysis.AutofilterHealth analysis) {
    return new AutofilterHealthReport(
        analysis.checkedAutofilterCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(WorkbookReadResultConverter::toAnalysisFindingReport)
            .toList());
  }

  private static TableHealthReport toTableHealthReport(WorkbookAnalysis.TableHealth analysis) {
    return new TableHealthReport(
        analysis.checkedTableCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(WorkbookReadResultConverter::toAnalysisFindingReport)
            .toList());
  }

  private static GridGrindResponse.HyperlinkHealthReport toHyperlinkHealthReport(
      WorkbookAnalysis.HyperlinkHealth analysis) {
    return new GridGrindResponse.HyperlinkHealthReport(
        analysis.checkedHyperlinkCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(WorkbookReadResultConverter::toAnalysisFindingReport)
            .toList());
  }

  private static GridGrindResponse.NamedRangeHealthReport toNamedRangeHealthReport(
      WorkbookAnalysis.NamedRangeHealth analysis) {
    return new GridGrindResponse.NamedRangeHealthReport(
        analysis.checkedNamedRangeCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(WorkbookReadResultConverter::toAnalysisFindingReport)
            .toList());
  }

  private static GridGrindResponse.WorkbookFindingsReport toWorkbookFindingsReport(
      WorkbookAnalysis.WorkbookFindings analysis) {
    return new GridGrindResponse.WorkbookFindingsReport(
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(WorkbookReadResultConverter::toAnalysisFindingReport)
            .toList());
  }

  private static GridGrindResponse.AnalysisSummaryReport toAnalysisSummaryReport(
      WorkbookAnalysis.AnalysisSummary summary) {
    return new GridGrindResponse.AnalysisSummaryReport(
        summary.totalCount(), summary.errorCount(), summary.warningCount(), summary.infoCount());
  }

  private static GridGrindResponse.AnalysisFindingReport toAnalysisFindingReport(
      WorkbookAnalysis.AnalysisFinding finding) {
    return new GridGrindResponse.AnalysisFindingReport(
        toAnalysisFindingCode(finding.code()),
        toAnalysisSeverity(finding.severity()),
        finding.title(),
        finding.message(),
        toAnalysisLocationReport(finding.location()),
        finding.evidence());
  }

  private static GridGrindResponse.AnalysisLocationReport toAnalysisLocationReport(
      WorkbookAnalysis.AnalysisLocation location) {
    return switch (location) {
      case WorkbookAnalysis.AnalysisLocation.Workbook _ ->
          new GridGrindResponse.AnalysisLocationReport.Workbook();
      case WorkbookAnalysis.AnalysisLocation.Sheet sheet ->
          new GridGrindResponse.AnalysisLocationReport.Sheet(sheet.sheetName());
      case WorkbookAnalysis.AnalysisLocation.Cell cell ->
          new GridGrindResponse.AnalysisLocationReport.Cell(cell.sheetName(), cell.address());
      case WorkbookAnalysis.AnalysisLocation.Range range ->
          new GridGrindResponse.AnalysisLocationReport.Range(range.sheetName(), range.range());
      case WorkbookAnalysis.AnalysisLocation.NamedRange namedRange ->
          new GridGrindResponse.AnalysisLocationReport.NamedRange(
              namedRange.name(), toNamedRangeScope(namedRange.scope()));
    };
  }

  private static AutofilterEntryReport toAutofilterEntryReport(ExcelAutofilterSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelAutofilterSnapshot.SheetOwned sheetOwned ->
          new AutofilterEntryReport.SheetOwned(
              sheetOwned.range(),
              sheetOwned.filterColumns().stream()
                  .map(WorkbookReadResultConverter::toAutofilterFilterColumnReport)
                  .toList(),
              toAutofilterSortStateReport(sheetOwned.sortState()));
      case ExcelAutofilterSnapshot.TableOwned tableOwned ->
          new AutofilterEntryReport.TableOwned(
              tableOwned.range(),
              tableOwned.tableName(),
              tableOwned.filterColumns().stream()
                  .map(WorkbookReadResultConverter::toAutofilterFilterColumnReport)
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
        snapshot.columns().stream().map(WorkbookReadResultConverter::toTableColumnReport).toList(),
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

  private static GridGrindResponse.CommentReport toCommentReport(ExcelCommentSnapshot comment) {
    if (comment == null) {
      return null;
    }
    return new GridGrindResponse.CommentReport(
        comment.text(),
        comment.author(),
        comment.visible(),
        toRichTextRunReports(comment.runs()),
        comment.anchor() == null
            ? null
            : new CommentAnchorReport(
                comment.anchor().firstColumn(),
                comment.anchor().firstRow(),
                comment.anchor().lastColumn(),
                comment.anchor().lastRow()));
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

  private static AnalysisFindingCode toAnalysisFindingCode(
      WorkbookAnalysis.AnalysisFindingCode code) {
    return AnalysisFindingCode.valueOf(code.name());
  }

  private static AnalysisSeverity toAnalysisSeverity(WorkbookAnalysis.AnalysisSeverity severity) {
    return AnalysisSeverity.valueOf(severity.name());
  }
}
