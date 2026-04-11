package dev.erst.gridgrind.protocol.exec;

import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;

/** Converts protocol write operations and style inputs into workbook-core commands. */
final class WorkbookCommandConverter {
  private WorkbookCommandConverter() {}

  /** Converts one protocol operation into the matching workbook-core command. */
  static WorkbookCommand toCommand(WorkbookOperation operation) {
    return switch (operation) {
      case WorkbookOperation.EnsureSheet op -> new WorkbookCommand.CreateSheet(op.sheetName());
      case WorkbookOperation.RenameSheet op ->
          new WorkbookCommand.RenameSheet(op.sheetName(), op.newSheetName());
      case WorkbookOperation.DeleteSheet op -> new WorkbookCommand.DeleteSheet(op.sheetName());
      case WorkbookOperation.MoveSheet op ->
          new WorkbookCommand.MoveSheet(op.sheetName(), op.targetIndex());
      case WorkbookOperation.CopySheet op ->
          new WorkbookCommand.CopySheet(
              op.sourceSheetName(), op.newSheetName(), toExcelSheetCopyPosition(op.position()));
      case WorkbookOperation.SetActiveSheet op ->
          new WorkbookCommand.SetActiveSheet(op.sheetName());
      case WorkbookOperation.SetSelectedSheets op ->
          new WorkbookCommand.SetSelectedSheets(op.sheetNames());
      case WorkbookOperation.SetSheetVisibility op ->
          new WorkbookCommand.SetSheetVisibility(op.sheetName(), op.visibility());
      case WorkbookOperation.SetSheetProtection op ->
          new WorkbookCommand.SetSheetProtection(
              op.sheetName(), toExcelSheetProtectionSettings(op.protection()), op.password());
      case WorkbookOperation.ClearSheetProtection op ->
          new WorkbookCommand.ClearSheetProtection(op.sheetName());
      case WorkbookOperation.SetWorkbookProtection op ->
          new WorkbookCommand.SetWorkbookProtection(
              toExcelWorkbookProtectionSettings(op.protection()));
      case WorkbookOperation.ClearWorkbookProtection _ ->
          new WorkbookCommand.ClearWorkbookProtection();
      case WorkbookOperation.MergeCells op ->
          new WorkbookCommand.MergeCells(op.sheetName(), op.range());
      case WorkbookOperation.UnmergeCells op ->
          new WorkbookCommand.UnmergeCells(op.sheetName(), op.range());
      case WorkbookOperation.SetColumnWidth op ->
          new WorkbookCommand.SetColumnWidth(
              op.sheetName(), op.firstColumnIndex(), op.lastColumnIndex(), op.widthCharacters());
      case WorkbookOperation.SetRowHeight op ->
          new WorkbookCommand.SetRowHeight(
              op.sheetName(), op.firstRowIndex(), op.lastRowIndex(), op.heightPoints());
      case WorkbookOperation.InsertRows op ->
          new WorkbookCommand.InsertRows(op.sheetName(), op.rowIndex(), op.rowCount());
      case WorkbookOperation.DeleteRows op ->
          new WorkbookCommand.DeleteRows(op.sheetName(), toExcelRowSpan(op.rows()));
      case WorkbookOperation.ShiftRows op ->
          new WorkbookCommand.ShiftRows(op.sheetName(), toExcelRowSpan(op.rows()), op.delta());
      case WorkbookOperation.InsertColumns op ->
          new WorkbookCommand.InsertColumns(op.sheetName(), op.columnIndex(), op.columnCount());
      case WorkbookOperation.DeleteColumns op ->
          new WorkbookCommand.DeleteColumns(op.sheetName(), toExcelColumnSpan(op.columns()));
      case WorkbookOperation.ShiftColumns op ->
          new WorkbookCommand.ShiftColumns(
              op.sheetName(), toExcelColumnSpan(op.columns()), op.delta());
      case WorkbookOperation.SetRowVisibility op ->
          new WorkbookCommand.SetRowVisibility(
              op.sheetName(), toExcelRowSpan(op.rows()), op.hidden());
      case WorkbookOperation.SetColumnVisibility op ->
          new WorkbookCommand.SetColumnVisibility(
              op.sheetName(), toExcelColumnSpan(op.columns()), op.hidden());
      case WorkbookOperation.GroupRows op ->
          new WorkbookCommand.GroupRows(op.sheetName(), toExcelRowSpan(op.rows()), op.collapsed());
      case WorkbookOperation.UngroupRows op ->
          new WorkbookCommand.UngroupRows(op.sheetName(), toExcelRowSpan(op.rows()));
      case WorkbookOperation.GroupColumns op ->
          new WorkbookCommand.GroupColumns(
              op.sheetName(), toExcelColumnSpan(op.columns()), op.collapsed());
      case WorkbookOperation.UngroupColumns op ->
          new WorkbookCommand.UngroupColumns(op.sheetName(), toExcelColumnSpan(op.columns()));
      case WorkbookOperation.SetSheetPane op ->
          new WorkbookCommand.SetSheetPane(op.sheetName(), toExcelSheetPane(op.pane()));
      case WorkbookOperation.SetSheetZoom op ->
          new WorkbookCommand.SetSheetZoom(op.sheetName(), op.zoomPercent());
      case WorkbookOperation.SetPrintLayout op ->
          new WorkbookCommand.SetPrintLayout(op.sheetName(), toExcelPrintLayout(op.printLayout()));
      case WorkbookOperation.ClearPrintLayout op ->
          new WorkbookCommand.ClearPrintLayout(op.sheetName());
      case WorkbookOperation.SetCell op ->
          new WorkbookCommand.SetCell(op.sheetName(), op.address(), toExcelCellValue(op.value()));
      case WorkbookOperation.SetRange op ->
          new WorkbookCommand.SetRange(
              op.sheetName(),
              op.range(),
              op.rows().stream()
                  .map(row -> row.stream().map(WorkbookCommandConverter::toExcelCellValue).toList())
                  .toList());
      case WorkbookOperation.ClearRange op ->
          new WorkbookCommand.ClearRange(op.sheetName(), op.range());
      case WorkbookOperation.SetHyperlink op ->
          new WorkbookCommand.SetHyperlink(
              op.sheetName(), op.address(), toExcelHyperlink(op.target()));
      case WorkbookOperation.ClearHyperlink op ->
          new WorkbookCommand.ClearHyperlink(op.sheetName(), op.address());
      case WorkbookOperation.SetComment op ->
          new WorkbookCommand.SetComment(
              op.sheetName(), op.address(), toExcelComment(op.comment()));
      case WorkbookOperation.ClearComment op ->
          new WorkbookCommand.ClearComment(op.sheetName(), op.address());
      case WorkbookOperation.ApplyStyle op ->
          new WorkbookCommand.ApplyStyle(op.sheetName(), op.range(), toExcelCellStyle(op.style()));
      case WorkbookOperation.SetDataValidation op ->
          new WorkbookCommand.SetDataValidation(
              op.sheetName(), op.range(), toExcelDataValidationDefinition(op.validation()));
      case WorkbookOperation.ClearDataValidations op ->
          new WorkbookCommand.ClearDataValidations(
              op.sheetName(), toExcelRangeSelection(op.selection()));
      case WorkbookOperation.SetConditionalFormatting op ->
          new WorkbookCommand.SetConditionalFormatting(
              op.sheetName(), toExcelConditionalFormattingBlock(op.conditionalFormatting()));
      case WorkbookOperation.ClearConditionalFormatting op ->
          new WorkbookCommand.ClearConditionalFormatting(
              op.sheetName(), toExcelRangeSelection(op.selection()));
      case WorkbookOperation.SetAutofilter op ->
          new WorkbookCommand.SetAutofilter(
              op.sheetName(),
              op.range(),
              op.criteria().stream()
                  .map(WorkbookCommandConverter::toExcelAutofilterFilterColumn)
                  .toList(),
              toExcelAutofilterSortState(op.sortState()));
      case WorkbookOperation.ClearAutofilter op ->
          new WorkbookCommand.ClearAutofilter(op.sheetName());
      case WorkbookOperation.SetTable op ->
          new WorkbookCommand.SetTable(toExcelTableDefinition(op.table()));
      case WorkbookOperation.DeleteTable op ->
          new WorkbookCommand.DeleteTable(op.name(), op.sheetName());
      case WorkbookOperation.SetNamedRange op ->
          new WorkbookCommand.SetNamedRange(
              new ExcelNamedRangeDefinition(
                  op.name(),
                  toExcelNamedRangeScope(op.scope()),
                  toExcelNamedRangeTarget(op.target())));
      case WorkbookOperation.DeleteNamedRange op ->
          new WorkbookCommand.DeleteNamedRange(op.name(), toExcelNamedRangeScope(op.scope()));
      case WorkbookOperation.AppendRow op ->
          new WorkbookCommand.AppendRow(
              op.sheetName(),
              op.values().stream().map(WorkbookCommandConverter::toExcelCellValue).toList());
      case WorkbookOperation.AutoSizeColumns op ->
          new WorkbookCommand.AutoSizeColumns(op.sheetName());
      case WorkbookOperation.EvaluateFormulas _ -> new WorkbookCommand.EvaluateAllFormulas();
      case WorkbookOperation.EvaluateFormulaCells op ->
          new WorkbookCommand.EvaluateFormulaCells(
              FormulaEnvironmentConverter.toExcelFormulaCellTargets(op.cells()));
      case WorkbookOperation.ClearFormulaCaches _ -> new WorkbookCommand.ClearFormulaCaches();
      case WorkbookOperation.ForceFormulaRecalculationOnOpen _ ->
          new WorkbookCommand.ForceFormulaRecalculationOnOpen();
    };
  }

  static ExcelCellValue toExcelCellValue(CellInput value) {
    return switch (value) {
      case CellInput.Blank _ -> ExcelCellValue.blank();
      case CellInput.Text text -> ExcelCellValue.text(text.text());
      case CellInput.RichText richText -> ExcelCellValue.richText(toExcelRichText(richText));
      case CellInput.Numeric numeric -> ExcelCellValue.number(numeric.number());
      case CellInput.BooleanValue booleanValue -> ExcelCellValue.bool(booleanValue.bool());
      case CellInput.Date date -> ExcelCellValue.date(date.date());
      case CellInput.DateTime dateTime -> ExcelCellValue.dateTime(dateTime.dateTime());
      case CellInput.Formula formula -> ExcelCellValue.formula(formula.formula());
    };
  }

  static ExcelRichText toExcelRichText(CellInput.RichText richText) {
    return new ExcelRichText(
        richText.runs().stream().map(WorkbookCommandConverter::toExcelRichTextRun).toList());
  }

  static ExcelRichTextRun toExcelRichTextRun(RichTextRunInput run) {
    return new ExcelRichTextRun(run.text(), toExcelCellFont(run.font()));
  }

  static ExcelHyperlink toExcelHyperlink(HyperlinkTarget target) {
    return switch (target) {
      case HyperlinkTarget.Url url -> new ExcelHyperlink.Url(url.target());
      case HyperlinkTarget.Email email -> new ExcelHyperlink.Email(email.email());
      case HyperlinkTarget.File file -> new ExcelHyperlink.File(file.path());
      case HyperlinkTarget.Document document -> new ExcelHyperlink.Document(document.target());
    };
  }

  static ExcelComment toExcelComment(CommentInput comment) {
    return new ExcelComment(
        comment.text(),
        comment.author(),
        comment.visible(),
        comment.runs() == null
            ? null
            : new ExcelRichText(
                comment.runs().stream().map(WorkbookCommandConverter::toExcelRichTextRun).toList()),
        comment.anchor() == null
            ? null
            : new ExcelCommentAnchor(
                comment.anchor().firstColumn(),
                comment.anchor().firstRow(),
                comment.anchor().lastColumn(),
                comment.anchor().lastRow()));
  }

  static ExcelRowSpan toExcelRowSpan(RowSpanInput rows) {
    return switch (rows) {
      case RowSpanInput.Band band -> new ExcelRowSpan(band.firstRowIndex(), band.lastRowIndex());
    };
  }

  static ExcelColumnSpan toExcelColumnSpan(ColumnSpanInput columns) {
    return switch (columns) {
      case ColumnSpanInput.Band band ->
          new ExcelColumnSpan(band.firstColumnIndex(), band.lastColumnIndex());
    };
  }

  static ExcelCellStyle toExcelCellStyle(CellStyleInput style) {
    return new ExcelCellStyle(
        style.numberFormat(),
        toExcelCellAlignment(style.alignment()),
        toExcelCellFont(style.font()),
        toExcelCellFill(style.fill()),
        toExcelBorder(style.border()),
        toExcelCellProtection(style.protection()));
  }

  static ExcelCellAlignment toExcelCellAlignment(CellAlignmentInput alignment) {
    if (alignment == null) {
      return null;
    }
    return new ExcelCellAlignment(
        alignment.wrapText(),
        alignment.horizontalAlignment(),
        alignment.verticalAlignment(),
        alignment.textRotation(),
        alignment.indentation());
  }

  static ExcelCellFont toExcelCellFont(CellFontInput font) {
    if (font == null) {
      return null;
    }
    return new ExcelCellFont(
        font.bold(),
        font.italic(),
        font.fontName(),
        toExcelFontHeight(font.fontHeight()),
        toExcelColor(
            font.fontColor(), font.fontColorTheme(), font.fontColorIndexed(), font.fontColorTint()),
        font.underline(),
        font.strikeout());
  }

  static ExcelCellFill toExcelCellFill(CellFillInput fill) {
    if (fill == null) {
      return null;
    }
    return new ExcelCellFill(
        fill.pattern(),
        toExcelColor(
            fill.foregroundColor(),
            fill.foregroundColorTheme(),
            fill.foregroundColorIndexed(),
            fill.foregroundColorTint()),
        toExcelColor(
            fill.backgroundColor(),
            fill.backgroundColorTheme(),
            fill.backgroundColorIndexed(),
            fill.backgroundColorTint()),
        toExcelGradientFill(fill.gradient()));
  }

  static ExcelCellProtection toExcelCellProtection(CellProtectionInput protection) {
    if (protection == null) {
      return null;
    }
    return new ExcelCellProtection(protection.locked(), protection.hiddenFormula());
  }

  static ExcelFontHeight toExcelFontHeight(FontHeightInput fontHeight) {
    if (fontHeight == null) {
      return null;
    }
    return switch (fontHeight) {
      case FontHeightInput.Points points -> ExcelFontHeight.fromPoints(points.points());
      case FontHeightInput.Twips twips -> new ExcelFontHeight(twips.twips());
    };
  }

  static ExcelBorder toExcelBorder(CellBorderInput border) {
    if (border == null) {
      return null;
    }
    return new ExcelBorder(
        toExcelBorderSide(border.all()),
        toExcelBorderSide(border.top()),
        toExcelBorderSide(border.right()),
        toExcelBorderSide(border.bottom()),
        toExcelBorderSide(border.left()));
  }

  static ExcelBorderSide toExcelBorderSide(CellBorderSideInput side) {
    return side == null
        ? null
        : new ExcelBorderSide(
            side.style(),
            toExcelColor(side.color(), side.colorTheme(), side.colorIndexed(), side.colorTint()));
  }

  static ExcelDataValidationDefinition toExcelDataValidationDefinition(
      DataValidationInput validation) {
    return new ExcelDataValidationDefinition(
        toExcelDataValidationRule(validation.rule()),
        validation.allowBlank(),
        validation.suppressDropDownArrow(),
        toExcelDataValidationPrompt(validation.prompt()),
        toExcelDataValidationErrorAlert(validation.errorAlert()));
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

  static ExcelDataValidationPrompt toExcelDataValidationPrompt(DataValidationPromptInput prompt) {
    return prompt == null
        ? null
        : new ExcelDataValidationPrompt(prompt.title(), prompt.text(), prompt.showPromptBox());
  }

  static ExcelDataValidationErrorAlert toExcelDataValidationErrorAlert(
      DataValidationErrorAlertInput errorAlert) {
    return errorAlert == null
        ? null
        : new ExcelDataValidationErrorAlert(
            errorAlert.style(), errorAlert.title(), errorAlert.text(), errorAlert.showErrorBox());
  }

  static ExcelConditionalFormattingBlockDefinition toExcelConditionalFormattingBlock(
      ConditionalFormattingBlockInput block) {
    return new ExcelConditionalFormattingBlockDefinition(
        block.ranges(),
        block.rules().stream()
            .map(WorkbookCommandConverter::toExcelConditionalFormattingRule)
            .toList());
  }

  static ExcelConditionalFormattingRule toExcelConditionalFormattingRule(
      ConditionalFormattingRuleInput rule) {
    return switch (rule) {
      case ConditionalFormattingRuleInput.FormulaRule formulaRule ->
          new ExcelConditionalFormattingRule.FormulaRule(
              formulaRule.formula(),
              formulaRule.stopIfTrue(),
              toExcelDifferentialStyle(formulaRule.style()));
      case ConditionalFormattingRuleInput.CellValueRule cellValueRule ->
          new ExcelConditionalFormattingRule.CellValueRule(
              cellValueRule.operator(),
              cellValueRule.formula1(),
              cellValueRule.formula2(),
              cellValueRule.stopIfTrue(),
              toExcelDifferentialStyle(cellValueRule.style()));
      case ConditionalFormattingRuleInput.ColorScaleRule colorScaleRule ->
          new ExcelConditionalFormattingRule.ColorScaleRule(
              colorScaleRule.thresholds().stream()
                  .map(WorkbookCommandConverter::toExcelConditionalFormattingThreshold)
                  .toList(),
              colorScaleRule.colors().stream().map(WorkbookCommandConverter::toExcelColor).toList(),
              colorScaleRule.stopIfTrue());
      case ConditionalFormattingRuleInput.DataBarRule dataBarRule ->
          new ExcelConditionalFormattingRule.DataBarRule(
              toExcelColor(dataBarRule.color()),
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
                  .map(WorkbookCommandConverter::toExcelConditionalFormattingThreshold)
                  .toList(),
              iconSetRule.stopIfTrue());
      case ConditionalFormattingRuleInput.Top10Rule top10Rule ->
          new ExcelConditionalFormattingRule.Top10Rule(
              top10Rule.rank(),
              top10Rule.percent(),
              top10Rule.bottom(),
              top10Rule.stopIfTrue(),
              toExcelDifferentialStyle(top10Rule.style()));
    };
  }

  static ExcelDifferentialStyle toExcelDifferentialStyle(DifferentialStyleInput style) {
    if (style == null) {
      return null;
    }
    return new ExcelDifferentialStyle(
        style.numberFormat(),
        style.bold(),
        style.italic(),
        toExcelFontHeight(style.fontHeight()),
        style.fontColor(),
        style.underline(),
        style.strikeout(),
        style.fillColor(),
        toExcelDifferentialBorder(style.border()));
  }

  static ExcelDifferentialBorder toExcelDifferentialBorder(DifferentialBorderInput border) {
    if (border == null) {
      return null;
    }
    return new ExcelDifferentialBorder(
        toExcelDifferentialBorderSide(border.all()),
        toExcelDifferentialBorderSide(border.top()),
        toExcelDifferentialBorderSide(border.right()),
        toExcelDifferentialBorderSide(border.bottom()),
        toExcelDifferentialBorderSide(border.left()));
  }

  private static ExcelDifferentialBorderSide toExcelDifferentialBorderSide(
      DifferentialBorderSideInput side) {
    return side == null ? null : new ExcelDifferentialBorderSide(side.style(), side.color());
  }

  static ExcelTableDefinition toExcelTableDefinition(TableInput table) {
    return new ExcelTableDefinition(
        table.name(),
        table.sheetName(),
        table.range(),
        table.showTotalsRow(),
        table.hasAutofilter(),
        toExcelTableStyle(table.style()),
        table.comment(),
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

  static ExcelNamedRangeTarget toExcelNamedRangeTarget(NamedRangeTarget target) {
    return target.formula() != null
        ? new ExcelNamedRangeTarget(target.formula())
        : new ExcelNamedRangeTarget(target.sheetName(), target.range());
  }

  /** Converts a protocol sheet copy-position variant into the engine position type. */
  static ExcelSheetCopyPosition toExcelSheetCopyPosition(SheetCopyPosition position) {
    return switch (position) {
      case SheetCopyPosition.AppendAtEnd _ -> new ExcelSheetCopyPosition.AppendAtEnd();
      case SheetCopyPosition.AtIndex atIndex ->
          new ExcelSheetCopyPosition.AtIndex(atIndex.targetIndex());
    };
  }

  private static ExcelSheetProtectionSettings toExcelSheetProtectionSettings(
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

  private static ExcelWorkbookProtectionSettings toExcelWorkbookProtectionSettings(
      WorkbookProtectionInput protection) {
    return new ExcelWorkbookProtectionSettings(
        protection.structureLocked(),
        protection.windowsLocked(),
        protection.revisionsLocked(),
        protection.workbookPassword(),
        protection.revisionsPassword());
  }

  private static ExcelSheetPane toExcelSheetPane(PaneInput pane) {
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

  private static ExcelPrintLayout toExcelPrintLayout(PrintLayoutInput printLayout) {
    return new ExcelPrintLayout(
        toExcelPrintArea(printLayout.printArea()),
        printLayout.orientation(),
        toExcelPrintScaling(printLayout.scaling()),
        toExcelPrintTitleRows(printLayout.repeatingRows()),
        toExcelPrintTitleColumns(printLayout.repeatingColumns()),
        new ExcelHeaderFooterText(
            printLayout.header().left(),
            printLayout.header().center(),
            printLayout.header().right()),
        new ExcelHeaderFooterText(
            printLayout.footer().left(),
            printLayout.footer().center(),
            printLayout.footer().right()),
        toExcelPrintSetup(printLayout.setup()));
  }

  private static ExcelColor toExcelColor(String rgb, Integer theme, Integer indexed, Double tint) {
    if (rgb == null && theme == null && indexed == null) {
      return null;
    }
    return new ExcelColor(rgb, theme, indexed, tint);
  }

  private static ExcelGradientFill toExcelGradientFill(CellGradientFillInput gradient) {
    if (gradient == null) {
      return null;
    }
    return new ExcelGradientFill(
        gradient.type(),
        gradient.degree(),
        gradient.left(),
        gradient.right(),
        gradient.top(),
        gradient.bottom(),
        gradient.stops().stream()
            .map(stop -> new ExcelGradientStop(stop.position(), toExcelColor(stop.color())))
            .toList());
  }

  private static ExcelColor toExcelColor(ColorInput color) {
    if (color == null) {
      return null;
    }
    return new ExcelColor(color.rgb(), color.theme(), color.indexed(), color.tint());
  }

  private static ExcelConditionalFormattingThreshold toExcelConditionalFormattingThreshold(
      ConditionalFormattingThresholdInput threshold) {
    return new ExcelConditionalFormattingThreshold(
        threshold.type(), threshold.formula(), threshold.value());
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

  private static ExcelAutofilterFilterColumn toExcelAutofilterFilterColumn(
      AutofilterFilterColumnInput column) {
    return new ExcelAutofilterFilterColumn(
        column.columnId(),
        column.showButton(),
        toExcelAutofilterFilterCriterion(column.criterion()));
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
          new ExcelAutofilterFilterCriterion.Color(color.cellColor(), toExcelColor(color.color()));
      case AutofilterFilterCriterionInput.Icon icon ->
          new ExcelAutofilterFilterCriterion.Icon(icon.iconSet(), icon.iconId());
    };
  }

  private static ExcelAutofilterSortState toExcelAutofilterSortState(
      AutofilterSortStateInput sortState) {
    if (sortState == null) {
      return null;
    }
    return new ExcelAutofilterSortState(
        sortState.range(),
        sortState.caseSensitive(),
        sortState.columnSort(),
        sortState.sortMethod(),
        sortState.conditions().stream()
            .map(WorkbookCommandConverter::toExcelAutofilterSortCondition)
            .toList());
  }

  private static ExcelAutofilterSortCondition toExcelAutofilterSortCondition(
      AutofilterSortConditionInput condition) {
    return new ExcelAutofilterSortCondition(
        condition.range(),
        condition.descending(),
        condition.sortBy(),
        toExcelColor(condition.color()),
        condition.iconId());
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

  private static ExcelRangeSelection toExcelRangeSelection(RangeSelection selection) {
    return switch (selection) {
      case RangeSelection.All _ -> new ExcelRangeSelection.All();
      case RangeSelection.Selected selected -> new ExcelRangeSelection.Selected(selected.ranges());
    };
  }
}
