package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetProtection;

/** Copies sheet-local workbook structures without relying on POI's raw sheet cloning. */
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
    retargetCopiedSheetFormulas(workbook, sourceSheetName, newSheetName, targetPoiSheet);
    copyComments(snapshot.comments(), targetSheet);
    copyMergedRegions(snapshot.mergedRegions(), targetSheet);
    copyLayout(snapshot.layout(), targetSheet);
    targetSheet.setPrintLayout(snapshot.printLayout());
    copyDataValidations(
        snapshot.dataValidations(), targetPoiSheet, workbook, sourceSheetName, newSheetName);
    copyConditionalFormatting(
        snapshot.conditionalFormattingBlocks(),
        targetSheet,
        workbook,
        sourceSheetName,
        newSheetName);
    copyTables(workbook, newSheetName, snapshot.tables());
    copyAutofilter(snapshot.sheetAutofilter(), targetSheet);
    copyLocalNamedRanges(workbook, sourceSheetName, newSheetName, snapshot.localNamedRanges());
    copyProtection(sourcePoiSheet, targetPoiSheet);

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
    ExcelSheet sourceSheet = workbook.sheet(sheetName);
    return new ExcelSheetCopySnapshot(
        copyableLocalNames(workbook.namedRanges(), sheetName),
        sourceSheet.layout(),
        sourceSheet.printLayout(),
        sourceSheet.mergedRegions(),
        sourceSheet.comments(new ExcelCellSelection.AllUsedCells()),
        copiedDataValidations(sourcePoiSheet),
        supportedConditionalFormatting(
            sourceSheet.conditionalFormatting(new ExcelRangeSelection.All()), sheetName),
        sheetOwnedAutofilter(autofilterController.sheetOwnedAutofilters(sourcePoiSheet)),
        tablesOnSheet(sourcePoiSheet));
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

  private static void retargetCopiedSheetFormulas(
      ExcelWorkbook workbook, String sourceSheetName, String newSheetName, XSSFSheet targetSheet) {
    int targetSheetIndex = workbook.xssfWorkbook().getSheetIndex(targetSheet);
    for (Row row : targetSheet) {
      for (Cell cell : row) {
        if (cell.getCellType() != CellType.FORMULA) {
          continue;
        }
        String rewritten =
            ExcelFormulaSheetRenameSupport.renameSheet(
                workbook.xssfWorkbook(),
                cell.getCellFormula(),
                FormulaType.CELL,
                targetSheetIndex,
                sourceSheetName,
                newSheetName);
        if (!rewritten.equals(cell.getCellFormula())) {
          cell.setCellFormula(rewritten);
        }
      }
    }
  }

  private static void copyComments(
      List<WorkbookReadResult.CellComment> comments, ExcelSheet targetSheet) {
    for (WorkbookReadResult.CellComment comment : comments) {
      targetSheet.setComment(comment.address(), comment.comment().toAuthoringComment());
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
    targetSheet.setPresentation(layout.presentation().toAuthoringPresentation());
  }

  private static void copyDataValidations(
      List<CTDataValidation> validations,
      XSSFSheet targetPoiSheet,
      ExcelWorkbook workbook,
      String sourceSheetName,
      String newSheetName) {
    if (validations.isEmpty()) {
      return;
    }
    var targetDataValidations = targetPoiSheet.getCTWorksheet().addNewDataValidations();
    int targetSheetIndex = workbook.xssfWorkbook().getSheetIndex(newSheetName);
    for (CTDataValidation validation : validations) {
      CTDataValidation copiedValidation = targetDataValidations.addNewDataValidation();
      copiedValidation.set(validation);
      retargetValidationFormulas(
          workbook, copiedValidation, targetSheetIndex, sourceSheetName, newSheetName);
    }
    targetDataValidations.setCount(targetDataValidations.sizeOfDataValidationArray());
  }

  private static void copyConditionalFormatting(
      List<ExcelConditionalFormattingBlockDefinition> blocks,
      ExcelSheet targetSheet,
      ExcelWorkbook workbook,
      String sourceSheetName,
      String newSheetName) {
    for (ExcelConditionalFormattingBlockDefinition block : blocks) {
      targetSheet.setConditionalFormatting(
          retargetConditionalFormattingBlock(workbook, block, sourceSheetName, newSheetName));
    }
  }

  private static void copyTables(
      ExcelWorkbook workbook, String targetSheetName, List<ExcelTableSnapshot> tables) {
    for (ExcelTableSnapshot table : tables) {
      workbook.setTable(copiedTableDefinition(workbook, targetSheetName, table));
    }
  }

  private static void copyAutofilter(
      ExcelAutofilterSnapshot.SheetOwned sheetAutofilter, ExcelSheet targetSheet) {
    if (sheetAutofilter == null) {
      return;
    }
    targetSheet.setAutofilter(
        sheetAutofilter.range(),
        sheetAutofilter.filterColumns().stream()
            .map(ExcelSheetCopyController::copyableAutofilterColumn)
            .toList(),
        copyableSortState(sheetAutofilter.sortState()));
  }

  private static void copyLocalNamedRanges(
      ExcelWorkbook workbook,
      String sourceSheetName,
      String targetSheetName,
      List<ExcelNamedRangeSnapshot> localNamedRanges) {
    ExcelNamedRangeScope.SheetScope scope = new ExcelNamedRangeScope.SheetScope(targetSheetName);
    int targetSheetIndex = workbook.xssfWorkbook().getSheetIndex(targetSheetName);
    for (ExcelNamedRangeSnapshot localNamedRange : localNamedRanges) {
      workbook.setNamedRange(
          copiedLocalNamedRange(
              workbook,
              localNamedRange,
              scope,
              targetSheetIndex,
              sourceSheetName,
              targetSheetName));
    }
  }

  private static void copyProtection(XSSFSheet sourceSheet, XSSFSheet targetSheet) {
    Objects.requireNonNull(sourceSheet, "sourceSheet must not be null");
    Objects.requireNonNull(targetSheet, "targetSheet must not be null");
    if (!sourceSheet.getProtect()) {
      return;
    }
    CTSheetProtection copiedProtection =
        (CTSheetProtection) sourceSheet.getCTWorksheet().getSheetProtection().copy();
    targetSheet.getCTWorksheet().setSheetProtection(copiedProtection);
  }

  /** Returns copyable sheet-local names scoped to the requested source sheet. */
  static List<ExcelNamedRangeSnapshot> copyableLocalNames(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    Objects.requireNonNull(namedRanges, "namedRanges must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
    List<ExcelNamedRangeSnapshot> localNamedRanges = new ArrayList<>();
    for (ExcelNamedRangeSnapshot namedRange : namedRanges) {
      switch (namedRange.scope()) {
        case ExcelNamedRangeScope.WorkbookScope _ -> {}
        case ExcelNamedRangeScope.SheetScope sheetScope -> {
          if (sheetScope.sheetName().equals(sheetName)) {
            localNamedRanges.add(namedRange);
          }
        }
      }
    }
    return List.copyOf(localNamedRanges);
  }

  /** Retained for focused tests covering explicit range-backed local names only. */
  static List<ExcelNamedRangeSnapshot.RangeSnapshot> copyableLocalRangeNames(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    return copyableLocalNames(namedRanges, sheetName).stream()
        .flatMap(
            namedRange ->
                namedRange instanceof ExcelNamedRangeSnapshot.RangeSnapshot rangeSnapshot
                    ? java.util.stream.Stream.of(rangeSnapshot)
                    : java.util.stream.Stream.empty())
        .toList();
  }

  /** Local names owned by a copied sheet are now fully copyable. */
  static void requireNoUncopyableLocalNamedRanges(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    Objects.requireNonNull(namedRanges, "namedRanges must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
  }

  /** Sheet copy now supports tables and no longer rejects them up front. */
  static void requireNoTables(XSSFSheet sheet, String sheetName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
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

  private static List<CTDataValidation> copiedDataValidations(XSSFSheet sourcePoiSheet) {
    if (!sourcePoiSheet.getCTWorksheet().isSetDataValidations()) {
      return List.of();
    }
    List<CTDataValidation> copiedValidations = new ArrayList<>();
    for (CTDataValidation validation :
        sourcePoiSheet.getCTWorksheet().getDataValidations().getDataValidationArray()) {
      copiedValidations.add((CTDataValidation) validation.copy());
    }
    return List.copyOf(copiedValidations);
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
      case ExcelConditionalFormattingRuleSnapshot.ColorScaleRule colorScaleRule ->
          new ExcelConditionalFormattingRule.ColorScaleRule(
              colorScaleRule.thresholds().stream()
                  .map(ExcelSheetCopyController::copyableThreshold)
                  .toList(),
              colorScaleRule.colors().stream().map(ExcelColor::new).toList(),
              colorScaleRule.stopIfTrue());
      case ExcelConditionalFormattingRuleSnapshot.DataBarRule dataBarRule ->
          new ExcelConditionalFormattingRule.DataBarRule(
              new ExcelColor(dataBarRule.color()),
              dataBarRule.iconOnly(),
              dataBarRule.widthMin(),
              dataBarRule.widthMax(),
              copyableThreshold(dataBarRule.minThreshold()),
              copyableThreshold(dataBarRule.maxThreshold()),
              dataBarRule.stopIfTrue());
      case ExcelConditionalFormattingRuleSnapshot.IconSetRule iconSetRule ->
          new ExcelConditionalFormattingRule.IconSetRule(
              iconSetRule.iconSet(),
              iconSetRule.iconOnly(),
              iconSetRule.reversed(),
              iconSetRule.thresholds().stream()
                  .map(ExcelSheetCopyController::copyableThreshold)
                  .toList(),
              iconSetRule.stopIfTrue());
      case ExcelConditionalFormattingRuleSnapshot.Top10Rule top10Rule ->
          new ExcelConditionalFormattingRule.Top10Rule(
              top10Rule.rank(),
              top10Rule.percent(),
              top10Rule.bottom(),
              top10Rule.stopIfTrue(),
              copyableStyle(top10Rule.style(), sourceSheetName));
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
    ExcelWorkbookSheetSupport.requireSheetName(sourceSheetName, "sourceSheetName");
    if (style == null) {
      return null;
    }
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

  /** Returns the one sheet-owned autofilter snapshot, or null when absent. */
  static ExcelAutofilterSnapshot.SheetOwned sheetOwnedAutofilter(
      List<ExcelAutofilterSnapshot> autofilters) {
    Objects.requireNonNull(autofilters, "autofilters must not be null");
    if (autofilters.isEmpty()) {
      return null;
    }
    ExcelAutofilterSnapshot autofilter = autofilters.getFirst();
    return switch (autofilter) {
      case ExcelAutofilterSnapshot.SheetOwned sheetOwned -> sheetOwned;
      case ExcelAutofilterSnapshot.TableOwned _ ->
          throw new IllegalStateException(
              "sheetOwnedAutofilters must not return table-owned autofilter snapshots");
    };
  }

  /** Retained for focused tests around the legacy range-only sheet autofilter path. */
  static Optional<String> sheetOwnedAutofilterRange(List<ExcelAutofilterSnapshot> autofilters) {
    ExcelAutofilterSnapshot.SheetOwned sheetOwned = sheetOwnedAutofilter(autofilters);
    return sheetOwned == null ? Optional.empty() : Optional.of(sheetOwned.range());
  }

  private static ExcelConditionalFormattingThreshold copyableThreshold(
      ExcelConditionalFormattingThresholdSnapshot threshold) {
    return new ExcelConditionalFormattingThreshold(
        threshold.type(), threshold.formula(), threshold.value());
  }

  private static ExcelAutofilterFilterColumn copyableAutofilterColumn(
      ExcelAutofilterFilterColumnSnapshot filterColumn) {
    return new ExcelAutofilterFilterColumn(
        filterColumn.columnId(),
        filterColumn.showButton(),
        copyableAutofilterCriterion(filterColumn.criterion()));
  }

  private static ExcelAutofilterFilterCriterion copyableAutofilterCriterion(
      ExcelAutofilterFilterCriterionSnapshot criterion) {
    return switch (criterion) {
      case ExcelAutofilterFilterCriterionSnapshot.Values values ->
          new ExcelAutofilterFilterCriterion.Values(values.values(), values.includeBlank());
      case ExcelAutofilterFilterCriterionSnapshot.Custom custom ->
          new ExcelAutofilterFilterCriterion.Custom(
              custom.and(),
              custom.conditions().stream()
                  .map(
                      condition ->
                          new ExcelAutofilterFilterCriterion.CustomCondition(
                              condition.operator(), condition.value()))
                  .toList());
      case ExcelAutofilterFilterCriterionSnapshot.Dynamic dynamic ->
          new ExcelAutofilterFilterCriterion.Dynamic(
              dynamic.type(), dynamic.value(), dynamic.maxValue());
      case ExcelAutofilterFilterCriterionSnapshot.Top10 top10 ->
          new ExcelAutofilterFilterCriterion.Top10(
              (int) Math.round(top10.value()), top10.top(), top10.percent());
      case ExcelAutofilterFilterCriterionSnapshot.Color color ->
          new ExcelAutofilterFilterCriterion.Color(
              color.cellColor(),
              new ExcelColor(
                  color.color().rgb(),
                  color.color().theme(),
                  color.color().indexed(),
                  color.color().tint()));
      case ExcelAutofilterFilterCriterionSnapshot.Icon icon ->
          new ExcelAutofilterFilterCriterion.Icon(icon.iconSet(), icon.iconId());
    };
  }

  private static ExcelAutofilterSortState copyableSortState(
      ExcelAutofilterSortStateSnapshot sortState) {
    if (sortState == null) {
      return null;
    }
    return new ExcelAutofilterSortState(
        sortState.range(),
        sortState.caseSensitive(),
        sortState.columnSort(),
        sortState.sortMethod(),
        sortState.conditions().stream()
            .map(ExcelSheetCopyController::copyableSortCondition)
            .toList());
  }

  private static ExcelAutofilterSortCondition copyableSortCondition(
      ExcelAutofilterSortConditionSnapshot condition) {
    return new ExcelAutofilterSortCondition(
        condition.range(),
        condition.descending(),
        condition.sortBy(),
        condition.color() == null
            ? null
            : new ExcelColor(
                condition.color().rgb(),
                condition.color().theme(),
                condition.color().indexed(),
                condition.color().tint()),
        condition.iconId());
  }

  private static List<ExcelTableSnapshot> tablesOnSheet(XSSFSheet sourcePoiSheet) {
    List<ExcelTableSnapshot> tables = new ArrayList<>();
    for (var table : sourcePoiSheet.getTables()) {
      tables.add(ExcelTableCatalogSupport.toSnapshot(sourcePoiSheet.getSheetName(), table));
    }
    return List.copyOf(tables);
  }

  private static ExcelTableDefinition copiedTableDefinition(
      ExcelWorkbook workbook, String targetSheetName, ExcelTableSnapshot table) {
    return new ExcelTableDefinition(
        uniqueCopiedTableName(workbook, table.name()),
        targetSheetName,
        table.range(),
        table.totalsRowCount() > 0,
        table.hasAutofilter(),
        switch (table.style()) {
          case ExcelTableStyleSnapshot.None _ -> new ExcelTableStyle.None();
          case ExcelTableStyleSnapshot.Named named ->
              new ExcelTableStyle.Named(
                  named.name(),
                  named.showFirstColumn(),
                  named.showLastColumn(),
                  named.showRowStripes(),
                  named.showColumnStripes());
        },
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
                        Math.toIntExact(column.id() - 1L),
                        column.uniqueName(),
                        column.totalsRowLabel(),
                        column.totalsRowFunction(),
                        column.calculatedColumnFormula()))
            .toList());
  }

  private static String uniqueCopiedTableName(ExcelWorkbook workbook, String baseName) {
    String candidate = baseName;
    int suffix = 2;
    while (tableNameExists(workbook, candidate)) {
      candidate = baseName + "_Copy" + suffix;
      suffix++;
    }
    return candidate;
  }

  private static boolean tableNameExists(ExcelWorkbook workbook, String candidate) {
    for (String sheetName : workbook.sheetNames()) {
      XSSFSheet sheet = ExcelWorkbookSheetSupport.requiredSheet(workbook.xssfWorkbook(), sheetName);
      for (var table : sheet.getTables()) {
        if (Objects.requireNonNullElse(table.getName(), "").equalsIgnoreCase(candidate)) {
          return true;
        }
      }
    }
    return false;
  }

  private static void retargetValidationFormulas(
      ExcelWorkbook workbook,
      CTDataValidation validation,
      int targetSheetIndex,
      String sourceSheetName,
      String newSheetName) {
    String type = validationType(validation);
    if ("list".equals(type)) {
      String formula1 = validation.isSetFormula1() ? validation.getFormula1() : "";
      if (shouldRetargetValidationListFormula(formula1)) {
        validation.setFormula1(
            retargetFormula(
                workbook,
                formula1,
                FormulaType.DATAVALIDATION_LIST,
                targetSheetIndex,
                sourceSheetName,
                newSheetName));
      }
      return;
    }
    if (validation.isSetFormula1() && !validation.getFormula1().isBlank()) {
      validation.setFormula1(
          retargetFormula(
              workbook,
              validation.getFormula1(),
              FormulaType.CELL,
              targetSheetIndex,
              sourceSheetName,
              newSheetName));
    }
    if (validation.isSetFormula2() && !validation.getFormula2().isBlank()) {
      validation.setFormula2(
          retargetFormula(
              workbook,
              validation.getFormula2(),
              FormulaType.CELL,
              targetSheetIndex,
              sourceSheetName,
              newSheetName));
    }
  }

  private static String validationType(CTDataValidation validation) {
    return validation.isSetType()
        ? validation.getType().toString().toLowerCase(java.util.Locale.ROOT)
        : "";
  }

  private static ExcelNamedRangeDefinition copiedLocalNamedRange(
      ExcelWorkbook workbook,
      ExcelNamedRangeSnapshot localNamedRange,
      ExcelNamedRangeScope.SheetScope scope,
      int targetSheetIndex,
      String sourceSheetName,
      String targetSheetName) {
    return switch (localNamedRange) {
      case ExcelNamedRangeSnapshot.RangeSnapshot rangeSnapshot ->
          new ExcelNamedRangeDefinition(
              rangeSnapshot.name(),
              scope,
              new ExcelNamedRangeTarget(targetSheetName, rangeSnapshot.target().range()));
      case ExcelNamedRangeSnapshot.FormulaSnapshot formulaSnapshot ->
          new ExcelNamedRangeDefinition(
              formulaSnapshot.name(),
              scope,
              new ExcelNamedRangeTarget(
                  ExcelFormulaSheetRenameSupport.renameSheet(
                      workbook.xssfWorkbook(),
                      formulaSnapshot.refersToFormula(),
                      FormulaType.NAMEDRANGE,
                      targetSheetIndex,
                      sourceSheetName,
                      targetSheetName)));
    };
  }

  private static boolean shouldRetargetValidationListFormula(String formula1) {
    return !formula1.isBlank() && !isQuotedListLiteral(formula1);
  }

  private static boolean isQuotedListLiteral(String formula) {
    return formula.length() >= 2 && formula.startsWith("\"") && formula.endsWith("\"");
  }

  private static ExcelConditionalFormattingBlockDefinition retargetConditionalFormattingBlock(
      ExcelWorkbook workbook,
      ExcelConditionalFormattingBlockDefinition block,
      String sourceSheetName,
      String newSheetName) {
    int targetSheetIndex = workbook.xssfWorkbook().getSheetIndex(newSheetName);
    return new ExcelConditionalFormattingBlockDefinition(
        block.ranges(),
        block.rules().stream()
            .map(
                rule ->
                    retargetConditionalFormattingRule(
                        workbook, rule, targetSheetIndex, sourceSheetName, newSheetName))
            .toList());
  }

  private static ExcelConditionalFormattingRule retargetConditionalFormattingRule(
      ExcelWorkbook workbook,
      ExcelConditionalFormattingRule rule,
      int targetSheetIndex,
      String sourceSheetName,
      String newSheetName) {
    return switch (rule) {
      case ExcelConditionalFormattingRule.FormulaRule formulaRule ->
          new ExcelConditionalFormattingRule.FormulaRule(
              retargetFormula(
                  workbook,
                  formulaRule.formula(),
                  FormulaType.CONDFORMAT,
                  targetSheetIndex,
                  sourceSheetName,
                  newSheetName),
              formulaRule.stopIfTrue(),
              formulaRule.style());
      case ExcelConditionalFormattingRule.CellValueRule cellValueRule ->
          new ExcelConditionalFormattingRule.CellValueRule(
              cellValueRule.operator(),
              retargetFormula(
                  workbook,
                  cellValueRule.formula1(),
                  FormulaType.CONDFORMAT,
                  targetSheetIndex,
                  sourceSheetName,
                  newSheetName),
              retargetOptionalFormula(
                  workbook,
                  cellValueRule.formula2(),
                  FormulaType.CONDFORMAT,
                  targetSheetIndex,
                  sourceSheetName,
                  newSheetName),
              cellValueRule.stopIfTrue(),
              cellValueRule.style());
      case ExcelConditionalFormattingRule.ColorScaleRule colorScaleRule -> colorScaleRule;
      case ExcelConditionalFormattingRule.DataBarRule dataBarRule -> dataBarRule;
      case ExcelConditionalFormattingRule.IconSetRule iconSetRule -> iconSetRule;
      case ExcelConditionalFormattingRule.Top10Rule top10Rule -> top10Rule;
    };
  }

  private static String retargetFormula(
      ExcelWorkbook workbook,
      String formula,
      FormulaType formulaType,
      int targetSheetIndex,
      String sourceSheetName,
      String newSheetName) {
    return ExcelFormulaSheetRenameSupport.renameSheet(
        workbook.xssfWorkbook(),
        formula,
        formulaType,
        targetSheetIndex,
        sourceSheetName,
        newSheetName);
  }

  private static String retargetOptionalFormula(
      ExcelWorkbook workbook,
      String formula,
      FormulaType formulaType,
      int targetSheetIndex,
      String sourceSheetName,
      String newSheetName) {
    return formula == null
        ? null
        : retargetFormula(
            workbook, formula, formulaType, targetSheetIndex, sourceSheetName, newSheetName);
  }

  /** Immutable copy plan for the sheet-local structures GridGrind can reproduce safely. */
  private record ExcelSheetCopySnapshot(
      List<ExcelNamedRangeSnapshot> localNamedRanges,
      WorkbookReadResult.SheetLayout layout,
      ExcelPrintLayout printLayout,
      List<WorkbookReadResult.MergedRegion> mergedRegions,
      List<WorkbookReadResult.CellComment> comments,
      List<CTDataValidation> dataValidations,
      List<ExcelConditionalFormattingBlockDefinition> conditionalFormattingBlocks,
      ExcelAutofilterSnapshot.SheetOwned sheetAutofilter,
      List<ExcelTableSnapshot> tables) {
    private ExcelSheetCopySnapshot {
      localNamedRanges = List.copyOf(localNamedRanges);
      Objects.requireNonNull(layout, "layout must not be null");
      Objects.requireNonNull(printLayout, "printLayout must not be null");
      mergedRegions = List.copyOf(mergedRegions);
      comments = List.copyOf(comments);
      dataValidations = List.copyOf(dataValidations);
      conditionalFormattingBlocks = List.copyOf(conditionalFormattingBlocks);
      tables = List.copyOf(tables);
    }
  }
}
