package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Copies supported sheet-local workbook structures without relying on POI's raw sheet cloning. */
final class ExcelSheetCopyController {
  private final ExcelAutofilterController autofilterController = new ExcelAutofilterController();

  /** Copies one sheet into a new visible, unselected sheet at the requested workbook position. */
  ExcelWorkbook copySheet(
      ExcelWorkbook workbook,
      String sourceSheetName,
      String newSheetName,
      ExcelSheetCopyPosition position) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sourceSheetName, "sourceSheetName");
    ExcelWorkbookSheetSupport.requireSheetName(newSheetName, "newSheetName");
    Objects.requireNonNull(position, "position must not be null");

    ExcelWorkbookSheetSupport.requireSheetNameAvailable(workbook.xssfWorkbook(), newSheetName, -1);

    XSSFSheet sourcePoiSheet =
        ExcelWorkbookSheetSupport.requiredSheet(workbook.xssfWorkbook(), sourceSheetName);
    ExcelSheetCopySnapshot snapshot = snapshot(workbook, sourceSheetName, sourcePoiSheet);

    XSSFSheet targetPoiSheet = workbook.xssfWorkbook().createSheet(newSheetName);
    ExcelSheet targetSheet = workbook.sheet(newSheetName);

    copyRows(sourcePoiSheet, targetPoiSheet);
    copyComments(snapshot.comments(), targetSheet);
    copyMergedRegions(snapshot.mergedRegions(), targetSheet);
    copyLayout(snapshot.layout(), targetSheet);
    targetSheet.setPrintLayout(snapshot.printLayout());
    copyDataValidations(snapshot.validations(), targetSheet);
    copyConditionalFormatting(snapshot.conditionalFormattingBlocks(), targetSheet);
    copyAutofilter(snapshot.sheetAutofilterRange(), targetSheet);
    copyLocalNamedRanges(workbook, newSheetName, snapshot.localNamedRanges());
    snapshot
        .protection()
        .ifPresent(protection -> workbook.setSheetProtection(newSheetName, protection));

    int targetSheetIndex = workbook.xssfWorkbook().getSheetIndex(targetPoiSheet);
    workbook.xssfWorkbook().setSheetVisibility(targetSheetIndex, SheetVisibility.VISIBLE);
    targetPoiSheet.setSelected(false);
    switch (position) {
      case ExcelSheetCopyPosition.AppendAtEnd _ -> {
        // createSheet already appends at the end.
      }
      case ExcelSheetCopyPosition.AtIndex atIndex -> {
        ExcelWorkbookSheetSupport.requireTargetIndex(
            workbook.xssfWorkbook(), atIndex.targetIndex());
        workbook.xssfWorkbook().setSheetOrder(newSheetName, atIndex.targetIndex());
      }
    }
    ExcelWorkbookSheetSupport.normalizeWorkbookViewState(workbook.xssfWorkbook());
    targetPoiSheet.setSelected(false);
    return workbook;
  }

  private ExcelSheetCopySnapshot snapshot(
      ExcelWorkbook workbook, String sheetName, XSSFSheet sourcePoiSheet) {
    List<ExcelNamedRangeSnapshot> namedRanges = workbook.namedRanges();
    requireNoUncopyableLocalNamedRanges(namedRanges, sheetName);
    requireNoTables(sourcePoiSheet, sheetName);

    ExcelSheet sourceSheet = workbook.sheet(sheetName);
    return new ExcelSheetCopySnapshot(
        copyableLocalRangeNames(namedRanges, sheetName),
        sourceSheet.layout(),
        sourceSheet.printLayout(),
        sourceSheet.mergedRegions(),
        sourceSheet.comments(new ExcelCellSelection.AllUsedCells()),
        supportedDataValidations(
            sourceSheet.dataValidations(new ExcelRangeSelection.All()), sheetName),
        supportedConditionalFormatting(
            sourceSheet.conditionalFormatting(new ExcelRangeSelection.All()), sheetName),
        sheetOwnedAutofilterRange(autofilterController.sheetOwnedAutofilters(sourcePoiSheet)),
        ExcelSheetProtectionSupport.settings(sourcePoiSheet));
  }

  private static void copyRows(XSSFSheet sourceSheet, XSSFSheet targetSheet) {
    List<Row> sourceRows = new ArrayList<>();
    for (Row sourceRow : sourceSheet) {
      sourceRows.add(sourceRow);
    }
    if (sourceRows.isEmpty()) {
      return;
    }

    CellCopyPolicy copyPolicy = new CellCopyPolicy();
    copyPolicy.setCopyCellValue(true);
    copyPolicy.setCopyCellStyle(true);
    copyPolicy.setCopyCellFormula(true);
    copyPolicy.setCopyHyperlink(true);
    copyPolicy.setMergeHyperlink(false);
    copyPolicy.setCopyRowHeight(true);
    copyPolicy.setCondenseRows(false);
    copyPolicy.setCopyMergedRegions(false);
    targetSheet.copyRows(sourceRows, 0, copyPolicy);
  }

  private static void copyComments(
      List<WorkbookReadResult.CellComment> comments, ExcelSheet targetSheet) {
    for (WorkbookReadResult.CellComment comment : comments) {
      targetSheet.setComment(comment.address(), comment.comment());
    }
  }

  private static void copyMergedRegions(
      List<WorkbookReadResult.MergedRegion> mergedRegions, ExcelSheet targetSheet) {
    for (WorkbookReadResult.MergedRegion mergedRegion : mergedRegions) {
      targetSheet.mergeCells(mergedRegion.range());
    }
  }

  private static void copyLayout(WorkbookReadResult.SheetLayout layout, ExcelSheet targetSheet) {
    for (WorkbookReadResult.ColumnLayout column : layout.columns()) {
      targetSheet.setColumnWidth(
          column.columnIndex(), column.columnIndex(), column.widthCharacters());
    }
    for (WorkbookReadResult.RowLayout row : layout.rows()) {
      targetSheet.setRowHeight(row.rowIndex(), row.rowIndex(), row.heightPoints());
    }
    targetSheet.setPane(layout.pane());
    targetSheet.setZoom(layout.zoomPercent());
  }

  private static void copyDataValidations(
      List<ExcelDataValidationSnapshot.Supported> validations, ExcelSheet targetSheet) {
    for (ExcelDataValidationSnapshot.Supported validation : validations) {
      for (String range : validation.ranges()) {
        targetSheet.setDataValidation(range, validation.validation());
      }
    }
  }

  private static void copyConditionalFormatting(
      List<ExcelConditionalFormattingBlockDefinition> blocks, ExcelSheet targetSheet) {
    for (ExcelConditionalFormattingBlockDefinition block : blocks) {
      targetSheet.setConditionalFormatting(block);
    }
  }

  private static void copyAutofilter(
      Optional<String> sheetAutofilterRange, ExcelSheet targetSheet) {
    sheetAutofilterRange.ifPresent(targetSheet::setAutofilter);
  }

  private static void copyLocalNamedRanges(
      ExcelWorkbook workbook,
      String targetSheetName,
      List<ExcelNamedRangeSnapshot.RangeSnapshot> localNamedRanges) {
    ExcelNamedRangeScope.SheetScope scope = new ExcelNamedRangeScope.SheetScope(targetSheetName);
    for (ExcelNamedRangeSnapshot.RangeSnapshot localNamedRange : localNamedRanges) {
      workbook.setNamedRange(localNamedRangeDefinition(targetSheetName, scope, localNamedRange));
    }
  }

  /** Returns copyable sheet-local range names scoped to the requested source sheet. */
  static List<ExcelNamedRangeSnapshot.RangeSnapshot> copyableLocalRangeNames(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    Objects.requireNonNull(namedRanges, "namedRanges must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
    List<ExcelNamedRangeSnapshot.RangeSnapshot> localNamedRanges = new ArrayList<>();
    for (ExcelNamedRangeSnapshot namedRange : namedRanges) {
      switch (namedRange) {
        case ExcelNamedRangeSnapshot.RangeSnapshot rangeSnapshot -> {
          switch (rangeSnapshot.scope()) {
            case ExcelNamedRangeScope.WorkbookScope _ -> {}
            case ExcelNamedRangeScope.SheetScope sheetScope -> {
              if (sheetScope.sheetName().equals(sheetName)) {
                localNamedRanges.add(rangeSnapshot);
              }
            }
          }
        }
        case ExcelNamedRangeSnapshot.FormulaSnapshot _ -> {
          // handled separately by requireNoUncopyableLocalNamedRanges
        }
      }
    }
    return List.copyOf(localNamedRanges);
  }

  /**
   * Fails when the requested sheet owns a formula-defined local named range that cannot be copied.
   */
  static void requireNoUncopyableLocalNamedRanges(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    Objects.requireNonNull(namedRanges, "namedRanges must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
    for (ExcelNamedRangeSnapshot namedRange : namedRanges) {
      switch (namedRange) {
        case ExcelNamedRangeSnapshot.RangeSnapshot _ -> {
          // copied separately after the destination sheet exists.
        }
        case ExcelNamedRangeSnapshot.FormulaSnapshot formulaSnapshot -> {
          switch (formulaSnapshot.scope()) {
            case ExcelNamedRangeScope.WorkbookScope _ -> {}
            case ExcelNamedRangeScope.SheetScope sheetScope -> {
              if (sheetScope.sheetName().equals(sheetName)) {
                throw new IllegalArgumentException(
                    "cannot copy sheet '"
                        + sheetName
                        + "': sheet-scoped formula-defined named ranges are not copyable");
              }
            }
          }
        }
      }
    }
  }

  /** Fails when the requested source sheet contains any table definitions. */
  static void requireNoTables(XSSFSheet sheet, String sheetName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
    if (!sheet.getTables().isEmpty()) {
      throw new IllegalArgumentException(
          "cannot copy sheet '" + sheetName + "': sheets containing tables are not copyable");
    }
  }

  /** Returns only copyable validation rules, rejecting unsupported rule families eagerly. */
  static List<ExcelDataValidationSnapshot.Supported> supportedDataValidations(
      List<ExcelDataValidationSnapshot> validations, String sourceSheetName) {
    Objects.requireNonNull(validations, "validations must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sourceSheetName, "sourceSheetName");
    List<ExcelDataValidationSnapshot.Supported> supportedValidations = new ArrayList<>();
    for (ExcelDataValidationSnapshot validation : validations) {
      switch (validation) {
        case ExcelDataValidationSnapshot.Supported supported -> supportedValidations.add(supported);
        case ExcelDataValidationSnapshot.Unsupported unsupported ->
            throw new IllegalArgumentException(
                "cannot copy sheet '"
                    + sourceSheetName
                    + "': unsupported data validation '"
                    + unsupported.kind()
                    + "' is not copyable");
      }
    }
    return List.copyOf(supportedValidations);
  }

  /** Returns only copyable conditional-formatting blocks for one source sheet. */
  static List<ExcelConditionalFormattingBlockDefinition> supportedConditionalFormatting(
      List<ExcelConditionalFormattingBlockSnapshot> blocks, String sourceSheetName) {
    Objects.requireNonNull(blocks, "blocks must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sourceSheetName, "sourceSheetName");
    List<ExcelConditionalFormattingBlockDefinition> copyableBlocks = new ArrayList<>();
    for (ExcelConditionalFormattingBlockSnapshot block : blocks) {
      copyableBlocks.add(
          copyableConditionalFormattingBlock(
              block, copyableConditionalFormattingRules(block.rules(), sourceSheetName)));
    }
    return List.copyOf(copyableBlocks);
  }

  private static ExcelNamedRangeDefinition localNamedRangeDefinition(
      String targetSheetName,
      ExcelNamedRangeScope.SheetScope scope,
      ExcelNamedRangeSnapshot.RangeSnapshot localNamedRange) {
    return new ExcelNamedRangeDefinition(
        localNamedRange.name(),
        scope,
        new ExcelNamedRangeTarget(targetSheetName, localNamedRange.target().range()));
  }

  private static ExcelConditionalFormattingBlockDefinition copyableConditionalFormattingBlock(
      ExcelConditionalFormattingBlockSnapshot block, List<ExcelConditionalFormattingRule> rules) {
    return new ExcelConditionalFormattingBlockDefinition(block.ranges(), List.copyOf(rules));
  }

  private static List<ExcelConditionalFormattingRule> copyableConditionalFormattingRules(
      List<ExcelConditionalFormattingRuleSnapshot> rules, String sourceSheetName) {
    List<ExcelConditionalFormattingRule> copyableRules = new ArrayList<>();
    for (ExcelConditionalFormattingRuleSnapshot rule : rules) {
      copyableRules.add(copyableRule(rule, sourceSheetName));
    }
    return List.copyOf(copyableRules);
  }

  /** Converts one persisted conditional-formatting rule into a copyable authoring rule. */
  static ExcelConditionalFormattingRule copyableRule(
      ExcelConditionalFormattingRuleSnapshot rule, String sourceSheetName) {
    Objects.requireNonNull(rule, "rule must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sourceSheetName, "sourceSheetName");
    return switch (rule) {
      case ExcelConditionalFormattingRuleSnapshot.FormulaRule formulaRule ->
          new ExcelConditionalFormattingRule.FormulaRule(
              formulaRule.formula(),
              formulaRule.stopIfTrue(),
              copyableStyle(formulaRule.style(), sourceSheetName));
      case ExcelConditionalFormattingRuleSnapshot.CellValueRule cellValueRule ->
          new ExcelConditionalFormattingRule.CellValueRule(
              cellValueRule.operator(),
              cellValueRule.formula1(),
              cellValueRule.formula2(),
              cellValueRule.stopIfTrue(),
              copyableStyle(cellValueRule.style(), sourceSheetName));
      case ExcelConditionalFormattingRuleSnapshot.ColorScaleRule _ ->
          throw new IllegalArgumentException(
              "cannot copy sheet '"
                  + sourceSheetName
                  + "': conditional-formatting color scales are not copyable");
      case ExcelConditionalFormattingRuleSnapshot.DataBarRule _ ->
          throw new IllegalArgumentException(
              "cannot copy sheet '"
                  + sourceSheetName
                  + "': conditional-formatting data bars are not copyable");
      case ExcelConditionalFormattingRuleSnapshot.IconSetRule _ ->
          throw new IllegalArgumentException(
              "cannot copy sheet '"
                  + sourceSheetName
                  + "': conditional-formatting icon sets are not copyable");
      case ExcelConditionalFormattingRuleSnapshot.UnsupportedRule unsupportedRule ->
          throw new IllegalArgumentException(
              "cannot copy sheet '"
                  + sourceSheetName
                  + "': unsupported conditional-formatting rule '"
                  + unsupportedRule.kind()
                  + "' is not copyable");
    };
  }

  /** Converts one supported differential style into its copyable authoring form. */
  static ExcelDifferentialStyle copyableStyle(
      ExcelDifferentialStyleSnapshot style, String sourceSheetName) {
    Objects.requireNonNull(style, "style must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sourceSheetName, "sourceSheetName");
    if (!style.unsupportedFeatures().isEmpty()) {
      throw new IllegalArgumentException(
          "cannot copy sheet '"
              + sourceSheetName
              + "': conditional-formatting rules with unsupported differential-style features are"
              + " not copyable");
    }
    return new ExcelDifferentialStyle(
        style.numberFormat(),
        style.bold(),
        style.italic(),
        style.fontHeight(),
        style.fontColor(),
        style.underline(),
        style.strikeout(),
        style.fillColor(),
        style.border());
  }

  /** Returns the one sheet-owned autofilter range, or fails if a table-owned snapshot leaks in. */
  static Optional<String> sheetOwnedAutofilterRange(List<ExcelAutofilterSnapshot> autofilters) {
    Objects.requireNonNull(autofilters, "autofilters must not be null");
    if (autofilters.isEmpty()) {
      return Optional.empty();
    }
    ExcelAutofilterSnapshot autofilter = autofilters.getFirst();
    return switch (autofilter) {
      case ExcelAutofilterSnapshot.SheetOwned sheetOwned -> Optional.of(sheetOwned.range());
      case ExcelAutofilterSnapshot.TableOwned _ ->
          throw new IllegalStateException(
              "sheetOwnedAutofilters must not return table-owned autofilter snapshots");
    };
  }

  /**
   * Immutable copy plan for the supported sheet-local structures GridGrind can reproduce safely.
   */
  private record ExcelSheetCopySnapshot(
      List<ExcelNamedRangeSnapshot.RangeSnapshot> localNamedRanges,
      WorkbookReadResult.SheetLayout layout,
      ExcelPrintLayout printLayout,
      List<WorkbookReadResult.MergedRegion> mergedRegions,
      List<WorkbookReadResult.CellComment> comments,
      List<ExcelDataValidationSnapshot.Supported> validations,
      List<ExcelConditionalFormattingBlockDefinition> conditionalFormattingBlocks,
      Optional<String> sheetAutofilterRange,
      Optional<ExcelSheetProtectionSettings> protection) {
    private ExcelSheetCopySnapshot {
      localNamedRanges = List.copyOf(localNamedRanges);
      Objects.requireNonNull(layout, "layout must not be null");
      Objects.requireNonNull(printLayout, "printLayout must not be null");
      mergedRegions = List.copyOf(mergedRegions);
      comments = List.copyOf(comments);
      validations = List.copyOf(validations);
      conditionalFormattingBlocks = List.copyOf(conditionalFormattingBlocks);
      Objects.requireNonNull(sheetAutofilterRange, "sheetAutofilterRange must not be null");
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }
}
