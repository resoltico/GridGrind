package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.style.DifferentialStyleReport;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableColumnSpec;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.autofilter.AutofilterAdapter;
import dev.erst.gridgrind.engine.charts.ChartAdapter;
import dev.erst.gridgrind.engine.charts.ChartAdapter.ResolvedAreaReference;
import dev.erst.gridgrind.engine.comments.CommentAdapter;
import dev.erst.gridgrind.engine.conditionalformatting.ConditionalFormattingAdapter;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import dev.erst.gridgrind.engine.formulas.FormulaSheetRename;
import dev.erst.gridgrind.engine.names.NamedRangeAdapter;
import dev.erst.gridgrind.engine.packaging.PackageRelations;
import dev.erst.gridgrind.engine.poi.FormulaWrites;
import dev.erst.gridgrind.engine.poi.ImageCatalog;
import dev.erst.gridgrind.engine.styles.ColorAdapter;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import org.apache.poi.ooxml.POIXMLDocumentPart.RelationPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetProtection;
import org.w3c.dom.Node;

/**
 * Owns sheet clone prerequisites, transient copy captures, formula retargeting and package repair.
 */
@SuppressWarnings("PMD.ExcessiveImports")
final class SheetCopyAdapter {

  private final AutofilterAdapter autofilterAdapter = new AutofilterAdapter();

  /** Copies one sheet into a new visible, unselected sheet at the requested workbook position. */
  XssfWorkbook copySheet(
      XssfWorkbook workbook,
      String sourceSheetName,
      String newSheetName,
      SheetCopyPosition position) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    SheetCollectionAdapter.requireSheetName(sourceSheetName, "sourceSheetName");
    SheetCollectionAdapter.requireSheetName(newSheetName, "newSheetName");
    Objects.requireNonNull(position, "position must not be null");

    SheetCollectionAdapter.requireSheetNameAvailable(workbook.xssfWorkbook(), newSheetName, -1);

    XSSFSheet sourcePoiSheet =
        SheetCollectionAdapter.requiredSheet(workbook.xssfWorkbook(), sourceSheetName);
    XssfSheet sourceSheet = workbook.sheet(sourceSheetName);
    PictureCopySnapshot pictures = snapshot(sourcePoiSheet);
    SheetCopyCapture snapshot = snapshot(workbook, sourceSheetName, sourcePoiSheet);
    EmbeddedObjectCopySnapshot embeddedObjects = snapshot(sourceSheet);
    RewrittenChartFormulas rewrittenChartFormulas = rewriteDefinedNameFormulas(sourcePoiSheet);
    prepareSourceSheetForClone(sourcePoiSheet);
    int sourceSheetIndex = workbook.xssfWorkbook().getSheetIndex(sourcePoiSheet);
    XSSFSheet targetPoiSheet;
    try {
      targetPoiSheet = workbook.xssfWorkbook().cloneSheet(sourceSheetIndex, newSheetName);
    } finally {
      rewrittenChartFormulas.restoreSourceFormulas();
    }
    XssfSheet targetSheet = workbook.sheet(newSheetName);
    List<String> transientCopiedTableNames =
        targetPoiSheet.getTables().stream()
            .map(table -> Objects.requireNonNullElse(table.getName(), ""))
            .toList();

    repairCopiedPictures(targetPoiSheet, pictures);
    retargetCopiedSheetFormulas(workbook, sourceSheetName, newSheetName, targetPoiSheet);
    replaceDataValidations(
        snapshot.dataValidations(), targetPoiSheet, workbook, sourceSheetName, newSheetName);
    replaceConditionalFormatting(
        snapshot.conditionalFormattingBlocks(),
        targetSheet,
        workbook,
        sourceSheetName,
        newSheetName);
    repairComments(snapshot.comments(), targetSheet);
    repairCopiedEmbeddedObjects(targetSheet, embeddedObjects);
    repairPrintLayout(snapshot.printLayout(), targetSheet);
    replaceTables(workbook, newSheetName, snapshot.tables());
    TableAdapter.retargetFormulaCells(
        targetPoiSheet,
        transientCopiedTableNames,
        targetPoiSheet.getTables().stream()
            .map(table -> Objects.requireNonNullElse(table.getName(), ""))
            .toList());
    TableAdapter.canonicalizeSheet(targetPoiSheet);
    replaceAutofilter(snapshot.sheetAutofilter(), targetSheet);
    replaceLocalNamedRanges(workbook, sourceSheetName, newSheetName, snapshot.localNamedRanges());
    copyProtection(sourcePoiSheet, targetPoiSheet);

    int targetSheetIndex = workbook.xssfWorkbook().getSheetIndex(targetPoiSheet);
    workbook.xssfWorkbook().setSheetVisibility(targetSheetIndex, SheetVisibility.VISIBLE);
    targetPoiSheet.setSelected(false);
    switch (position) {
      case SheetCopyPosition.AppendAtEnd _ -> {
        // cloneSheet already appends at the end.
      }
      case SheetCopyPosition.AtIndex atIndex -> {
        SheetCollectionAdapter.requireTargetIndex(workbook.xssfWorkbook(), atIndex.targetIndex());
        workbook.xssfWorkbook().setSheetOrder(newSheetName, atIndex.targetIndex());
      }
    }
    SheetCollectionAdapter.normalizeWorkbookViewState(workbook.xssfWorkbook());
    targetPoiSheet.setSelected(false);
    return workbook;
  }

  private SheetCopyCapture snapshot(
      XssfWorkbook workbook, String sheetName, XSSFSheet sourcePoiSheet) {
    XssfSheet sourceSheet = workbook.sheet(sheetName);
    return new SheetCopyCapture(
        copyableLocalNames(workbook.names().namedRanges(), sheetName),
        sourceSheet.layout().printLayout(),
        sourceSheet.comments().comments(),
        copiedDataValidations(sourcePoiSheet),
        snapshot(sourcePoiSheet, sheetName),
        sheetOwnedAutofilter(autofilterAdapter.sheetOwnedAutofilters(sourcePoiSheet)),
        tablesOnSheet(sourcePoiSheet));
  }

  static void retargetCopiedSheetFormulas(
      XssfWorkbook workbook, String sourceSheetName, String newSheetName, XSSFSheet targetSheet) {
    int targetSheetIndex = workbook.xssfWorkbook().getSheetIndex(targetSheet);
    for (Row row : targetSheet) {
      for (Cell cell : row) {
        if (cell.getCellType() != CellType.FORMULA) {
          continue;
        }
        if (!mayReferenceCopiedSheet(cell.getCellFormula(), sourceSheetName)) {
          continue;
        }
        String rewritten =
            FormulaSheetRename.renameSheet(
                workbook.xssfWorkbook(),
                cell.getCellFormula(),
                FormulaType.CELL,
                targetSheetIndex,
                sourceSheetName,
                newSheetName);
        if (!rewritten.equals(cell.getCellFormula())) {
          FormulaWrites.setRewrittenFormula(cell, rewritten, "Copied-sheet formula retargeting");
        }
      }
    }
  }

  static boolean mayReferenceCopiedSheet(String formula, String sourceSheetName) {
    String quotedSourceSheetName = "'" + sourceSheetName.replace("'", "''") + "'";
    return formula.contains(sourceSheetName + "!")
        || formula.contains(quotedSourceSheetName + "!")
        || formula.contains(sourceSheetName + ":")
        || formula.contains(quotedSourceSheetName + ":");
  }

  /** Immutable copy plan for the sheet-local structures GridGrind can reproduce safely. */
  private record SheetCopyCapture(
      List<NamedRangeReport> localNamedRanges,
      ResolvedPrintLayout printLayout,
      List<CellCommentReport> comments,
      List<CTDataValidation> dataValidations,
      List<org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting>
          conditionalFormattingBlocks,
      Optional<AutofilterEntryReport.SheetOwned> sheetAutofilter,
      List<TableEntryReport> tables) {
    private SheetCopyCapture {
      localNamedRanges = List.copyOf(localNamedRanges);
      Objects.requireNonNull(printLayout, "printLayout must not be null");
      comments = List.copyOf(comments);
      dataValidations = List.copyOf(dataValidations);
      conditionalFormattingBlocks = List.copyOf(conditionalFormattingBlocks);
      Objects.requireNonNull(sheetAutofilter, "sheetAutofilter must not be null");
      tables = List.copyOf(tables);
    }
  }

  static void repairComments(List<CellCommentReport> expectedComments, XssfSheet targetSheet) {
    if (expectedComments.isEmpty()) {
      return;
    }
    // cloneSheet can leave copied comments looking correct in-memory while reopening later
    // without a stable client anchor. Rewriting the copied comments makes the persisted
    // VML-backed comment state authoritative again.
    CommentAdapter commentRepairSupport = new CommentAdapter(targetSheet.xssfSheet());
    commentRepairSupport.replaceComments(CommentAdapter.commentRewriteSnapshots(expectedComments));
  }

  static void repairPrintLayout(ResolvedPrintLayout expectedPrintLayout, XssfSheet targetSheet) {
    if (expectedPrintLayout.equals(targetSheet.layout().printLayout())) {
      return;
    }
    targetSheet.layout().setPrintLayout(expectedPrintLayout);
  }

  static void copyProtection(XSSFSheet sourceSheet, XSSFSheet targetSheet) {
    Objects.requireNonNull(sourceSheet, "sourceSheet must not be null");
    Objects.requireNonNull(targetSheet, "targetSheet must not be null");
    if (!sourceSheet.getProtect()) {
      return;
    }
    CTSheetProtection copiedProtection =
        (CTSheetProtection) sourceSheet.getCTWorksheet().getSheetProtection().copy();
    targetSheet.getCTWorksheet().setSheetProtection(copiedProtection);
  }

  static String retargetFormula(
      XssfWorkbook workbook,
      String formula,
      FormulaType formulaType,
      int targetSheetIndex,
      String sourceSheetName,
      String newSheetName) {
    return FormulaSheetRename.renameSheet(
        workbook.xssfWorkbook(),
        formula,
        formulaType,
        targetSheetIndex,
        sourceSheetName,
        newSheetName);
  }

  static Optional<String> retargetOptionalFormula(
      XssfWorkbook workbook,
      Optional<String> formula,
      FormulaType formulaType,
      int targetSheetIndex,
      String sourceSheetName,
      String newSheetName) {
    Objects.requireNonNull(formula, "formula must not be null");
    return formula.map(
        value ->
            retargetFormula(
                workbook, value, formulaType, targetSheetIndex, sourceSheetName, newSheetName));
  }

  static void replaceAutofilter(
      Optional<AutofilterEntryReport.SheetOwned> sheetAutofilter, XssfSheet targetSheet) {
    targetSheet.metadata().clearAutofilter();
    sheetAutofilter.ifPresent(value -> copyAutofilter(value, targetSheet));
  }

  static Optional<AutofilterEntryReport.SheetOwned> sheetOwnedAutofilter(
      List<AutofilterEntryReport> autofilters) {
    Objects.requireNonNull(autofilters, "autofilters must not be null");
    if (autofilters.isEmpty()) {
      return Optional.empty();
    }
    AutofilterEntryReport autofilter = autofilters.getFirst();
    return switch (autofilter) {
      case AutofilterEntryReport.SheetOwned sheetOwned -> Optional.of(sheetOwned);
      case AutofilterEntryReport.TableOwned _ ->
          throw new IllegalStateException(
              "sheetOwnedAutofilters must not return table-owned autofilter snapshots");
    };
  }

  static AutofilterSortConditionSpec copyableSortCondition(
      AutofilterSortConditionReport condition) {
    return switch (condition) {
      case AutofilterSortConditionReport.Value value ->
          new AutofilterSortConditionSpec.Value(
              new dev.erst.gridgrind.core.vocabulary.RangeRef(value.range().value()),
              value.descending());
      case AutofilterSortConditionReport.CellColor cellColor ->
          new AutofilterSortConditionSpec.CellColor(
              new dev.erst.gridgrind.core.vocabulary.RangeRef(cellColor.range().value()),
              cellColor.descending(),
              Objects.requireNonNull(
                  ColorAdapter.copyOf(cellColor.color()), "cell sort color must not be null"));
      case AutofilterSortConditionReport.FontColor fontColor ->
          new AutofilterSortConditionSpec.FontColor(
              new dev.erst.gridgrind.core.vocabulary.RangeRef(fontColor.range().value()),
              fontColor.descending(),
              Objects.requireNonNull(
                  ColorAdapter.copyOf(fontColor.color()), "font sort color must not be null"));
      case AutofilterSortConditionReport.Icon icon ->
          new AutofilterSortConditionSpec.Icon(
              new dev.erst.gridgrind.core.vocabulary.RangeRef(icon.range().value()),
              icon.descending(),
              icon.iconId());
    };
  }

  private static void copyAutofilter(
      AutofilterEntryReport.SheetOwned sheetAutofilter, XssfSheet targetSheet) {
    targetSheet
        .metadata()
        .setAutofilter(
            sheetAutofilter.range().value(),
            sheetAutofilter.filterColumns().stream()
                .map(SheetCopyAdapter::copyableAutofilterColumn)
                .toList(),
            copyableSortState(sheetAutofilter.sortState()));
  }

  private static AutofilterFilterColumnSpec copyableAutofilterColumn(
      AutofilterFilterColumnReport filterColumn) {
    return new AutofilterFilterColumnSpec(
        filterColumn.columnId(),
        filterColumn.showButton(),
        copyableAutofilterCriterion(filterColumn.criterion()));
  }

  private static AutofilterFilterCriterionSpec copyableAutofilterCriterion(
      AutofilterFilterCriterionReport criterion) {
    return switch (criterion) {
      case AutofilterFilterCriterionReport.Values values ->
          new AutofilterFilterCriterionSpec.Values(values.values(), values.includeBlank());
      case AutofilterFilterCriterionReport.Custom custom ->
          new AutofilterFilterCriterionSpec.Custom(
              custom.and(),
              custom.conditions().stream()
                  .map(
                      condition ->
                          new AutofilterFilterCriterionSpec.CustomConditionSpec(
                              condition.operator(), condition.value()))
                  .toList());
      case AutofilterFilterCriterionReport.Dynamic dynamic ->
          new AutofilterFilterCriterionSpec.Dynamic(
              dynamic.dynamicType(),
              Optional.ofNullable(dynamic.value()),
              Optional.ofNullable(dynamic.maxValue()));
      case AutofilterFilterCriterionReport.Top10 top10 ->
          new AutofilterFilterCriterionSpec.Top10(
              (int) Math.round(top10.value()), top10.top(), top10.percent());
      case AutofilterFilterCriterionReport.Color color ->
          new AutofilterFilterCriterionSpec.Color(
              color.cellColor(),
              Objects.requireNonNull(
                  ColorAdapter.copyOf(color.color()), "autofilter color must not be null"));
      case AutofilterFilterCriterionReport.Icon icon ->
          new AutofilterFilterCriterionSpec.Icon(icon.iconSet(), icon.iconId());
    };
  }

  private static Optional<AutofilterSortStateSpec> copyableSortState(
      Optional<AutofilterSortStateReport> sortState) {
    return Objects.requireNonNull(sortState, "sortState must not be null")
        .map(
            snapshot ->
                new AutofilterSortStateSpec(
                    new dev.erst.gridgrind.core.vocabulary.RangeRef(snapshot.range().value()),
                    snapshot.caseSensitive(),
                    snapshot.columnSort(),
                    snapshot.sortMethod(),
                    snapshot.conditions().stream()
                        .map(SheetCopyAdapter::copyableSortCondition)
                        .toList()));
  }

  static List<CTConditionalFormatting> snapshot(XSSFSheet sheet, String sheetName) {
    var blocks = new ConditionalFormattingAdapter().conditionalFormatting(sheet);
    for (var block : blocks) {
      for (var rule : block.rules()) {
        switch (rule) {
          case ConditionalFormattingRuleReport.UnsupportedRule unsupported ->
              throw new IllegalArgumentException(
                  "cannot copy sheet '"
                      + sheetName
                      + "': unsupported conditional-formatting rule '"
                      + unsupported.kind()
                      + "' is not copyable");
          case ConditionalFormattingRuleReport.FormulaRule value ->
              requireSupportedStyle(value.style().orElse(null), sheetName);
          case ConditionalFormattingRuleReport.CellValueRule value ->
              requireSupportedStyle(value.style().orElse(null), sheetName);
          case ConditionalFormattingRuleReport.Top10Rule value ->
              requireSupportedStyle(value.style().orElse(null), sheetName);
          default -> {}
        }
      }
    }
    return sheet.getCTWorksheet().getConditionalFormattingList().stream()
        .map(block -> (CTConditionalFormatting) block.copy())
        .toList();
  }

  private static void requireSupportedStyle(
      @org.jspecify.annotations.Nullable DifferentialStyleReport style, String sheetName) {
    if (style != null && !style.unsupportedFeatures().isEmpty()) {
      throw new IllegalArgumentException(
          "cannot copy sheet '"
              + sheetName
              + "': unsupported conditional-formatting differential-style features "
              + style.unsupportedFeatures());
    }
  }

  static void replaceConditionalFormatting(
      List<CTConditionalFormatting> blocks,
      XssfSheet targetSheet,
      XssfWorkbook workbook,
      String sourceSheetName,
      String newSheetName) {
    var target = workbook.xssfWorkbook().getSheet(newSheetName);
    target.getCTWorksheet().setConditionalFormattingArray(new CTConditionalFormatting[0]);
    int index = workbook.xssfWorkbook().getSheetIndex(target);
    for (var original : blocks) {
      var block = target.getCTWorksheet().addNewConditionalFormatting();
      block.set(original);
      for (var rule : block.getCfRuleList()) {
        for (var threshold : thresholds(rule)) {
          if (threshold.getType()
                  == org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfvoType.FORMULA
              && threshold.getVal() != null) {
            threshold.setVal(
                retargetFormula(
                    workbook,
                    threshold.getVal(),
                    FormulaType.CONDFORMAT,
                    index,
                    sourceSheetName,
                    newSheetName));
          }
        }
        for (int formula = 0; formula < rule.sizeOfFormulaArray(); formula++) {
          rule.setFormulaArray(
              formula,
              retargetFormula(
                  workbook,
                  rule.getFormulaArray(formula),
                  FormulaType.CONDFORMAT,
                  index,
                  sourceSheetName,
                  newSheetName));
        }
      }
    }
  }

  private static List<org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfvo> thresholds(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule rule) {
    if (rule.isSetColorScale()) return rule.getColorScale().getCfvoList();
    if (rule.isSetDataBar()) return rule.getDataBar().getCfvoList();
    if (rule.isSetIconSet()) return rule.getIconSet().getCfvoList();
    return List.of();
  }

  static void replaceLocalNamedRanges(
      XssfWorkbook workbook,
      String sourceSheetName,
      String targetSheetName,
      List<NamedRangeReport> localNamedRanges) {
    deleteLocalNamedRanges(workbook, targetSheetName);
    copyLocalNamedRanges(workbook, sourceSheetName, targetSheetName, localNamedRanges);
  }

  static void deleteLocalNamedRanges(XssfWorkbook workbook, String sheetName) {
    NamedRangeScopeReport.Sheet scope =
        new NamedRangeScopeReport.Sheet(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName));
    for (NamedRangeReport localName :
        copyableLocalNames(workbook.names().namedRanges(), sheetName)) {
      workbook.names().deleteNamedRange(localName.name().value(), scope);
    }
  }

  static List<NamedRangeReport> copyableLocalNames(
      List<NamedRangeReport> namedRanges, String sheetName) {
    Objects.requireNonNull(namedRanges, "namedRanges must not be null");
    SheetCollectionAdapter.requireSheetName(sheetName, "sheetName");
    List<NamedRangeReport> localNamedRanges = new ArrayList<>();
    for (NamedRangeReport namedRange : namedRanges) {
      switch (namedRange.scope()) {
        case NamedRangeScopeReport.Workbook _ -> {}
        case NamedRangeScopeReport.Sheet sheetScope -> {
          if (sheetScope.sheetName().value().equals(sheetName)) {
            localNamedRanges.add(namedRange);
          }
        }
      }
    }
    return List.copyOf(localNamedRanges);
  }

  private static void copyLocalNamedRanges(
      XssfWorkbook workbook,
      String sourceSheetName,
      String targetSheetName,
      List<NamedRangeReport> localNamedRanges) {
    int targetSheetIndex = workbook.xssfWorkbook().getSheetIndex(targetSheetName);
    var destination = new dev.erst.gridgrind.core.vocabulary.SheetName(targetSheetName);
    for (NamedRangeReport localName : localNamedRanges) {
      String formula =
          switch (localName) {
            case NamedRangeReport.RangeReport range ->
                NamedRangeAdapter.rangeFormula(
                    destination, ((NamedRangeTarget.Range) range.target()).range());
            case NamedRangeReport.FormulaReport value ->
                value.refersToFormula().isEmpty()
                    ? ""
                    : retargetFormula(
                        workbook,
                        value.refersToFormula(),
                        FormulaType.NAMEDRANGE,
                        targetSheetIndex,
                        sourceSheetName,
                        targetSheetName);
          };
      var copied = workbook.xssfWorkbook().createName();
      copied.setNameName(localName.name().value());
      copied.setSheetIndex(targetSheetIndex);
      if (!formula.isEmpty()) copied.setRefersToFormula(formula);
    }
  }

  static void replaceDataValidations(
      List<CTDataValidation> validations,
      XSSFSheet targetPoiSheet,
      XssfWorkbook workbook,
      String sourceSheetName,
      String newSheetName) {
    if (targetPoiSheet.getCTWorksheet().isSetDataValidations()) {
      targetPoiSheet.getCTWorksheet().unsetDataValidations();
    }
    copyDataValidations(validations, targetPoiSheet, workbook, sourceSheetName, newSheetName);
  }

  static List<CTDataValidation> copiedDataValidations(XSSFSheet sourcePoiSheet) {
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

  private static void copyDataValidations(
      List<CTDataValidation> validations,
      XSSFSheet targetPoiSheet,
      XssfWorkbook workbook,
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

  private static void retargetValidationFormulas(
      XssfWorkbook workbook,
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
    return validation.isSetType() ? validation.getType().toString().toLowerCase(Locale.ROOT) : "";
  }

  private static boolean shouldRetargetValidationListFormula(String formula1) {
    return !formula1.isBlank() && !isQuotedListLiteral(formula1);
  }

  private static boolean isQuotedListLiteral(String formula) {
    return formula.length() >= 2 && formula.startsWith("\"") && formula.endsWith("\"");
  }

  static void replaceTables(
      XssfWorkbook workbook, String targetSheetName, List<TableEntryReport> tables) {
    List<String> existingTableNames =
        workbook.sheet(targetSheetName).xssfSheet().getTables().stream()
            .map(table -> table.getName())
            .toList();
    TableAdapter tableAdapter = new TableAdapter();
    for (String existingTableName : existingTableNames) {
      tableAdapter.deleteTable(workbook, existingTableName, targetSheetName);
    }
    copyTables(workbook, targetSheetName, tables);
  }

  static List<TableEntryReport> tablesOnSheet(XSSFSheet sourcePoiSheet) {
    List<TableEntryReport> tables = new ArrayList<>();
    for (var table : sourcePoiSheet.getTables()) {
      tables.add(TableAdapter.toSnapshot(sourcePoiSheet.getSheetName(), table));
    }
    return List.copyOf(tables);
  }

  private static void copyTables(
      XssfWorkbook workbook, String targetSheetName, List<TableEntryReport> tables) {
    TableAdapter tableAdapter = new TableAdapter();
    for (TableEntryReport table : tables) {
      tableAdapter.setTable(workbook, copiedTableDefinition(workbook, targetSheetName, table));
    }
  }

  private static ResolvedTable copiedTableDefinition(
      XssfWorkbook workbook, String targetSheetName, TableEntryReport table) {
    return new ResolvedTable(
        new TableName(uniqueCopiedTableName(workbook, table.name().value())),
        new SheetName(targetSheetName),
        new RangeRef(table.range().value()),
        table.structure().totalsRowCount() > 0,
        table.behavior().hasAutofilter(),
        switch (table.style()) {
          case TableStyleReport.None _ -> new TableStyleSpec.None();
          case TableStyleReport.Named named ->
              new TableStyleSpec.Named(
                  named.name(),
                  named.showFirstColumn(),
                  named.showLastColumn(),
                  named.showRowStripes(),
                  named.showColumnStripes());
        },
        table.presentation().comment().orElse(""),
        table.behavior().published(),
        table.behavior().insertRow(),
        table.behavior().insertRowShift(),
        table.presentation().headerRowCellStyle().orElse(""),
        table.presentation().dataCellStyle().orElse(""),
        table.presentation().totalsRowCellStyle().orElse(""),
        table.structure().columns().stream()
            .map(
                column ->
                    new TableColumnSpec(
                        Math.toIntExact(column.id() - 1L),
                        column.uniqueName().orElse(""),
                        column.totalsRowLabel().orElse(""),
                        column.totalsRowFunction().orElse(""),
                        column.calculatedColumnFormula().orElse("")))
            .toList());
  }

  private static String uniqueCopiedTableName(XssfWorkbook workbook, String baseName) {
    String candidate = baseName;
    int suffix = 2;
    while (tableNameExists(workbook, candidate)) {
      candidate = baseName + "_Copy" + suffix;
      suffix++;
    }
    return candidate;
  }

  private static boolean tableNameExists(XssfWorkbook workbook, String candidate) {
    for (String sheetName : workbook.sheets().sheetNames()) {
      XSSFSheet sheet = SheetCollectionAdapter.requiredSheet(workbook.xssfWorkbook(), sheetName);
      for (var table : sheet.getTables()) {
        if (Objects.requireNonNullElse(table.getName(), "").equalsIgnoreCase(candidate)) {
          return true;
        }
      }
    }
    return false;
  }

  PictureCopySnapshot snapshot(XSSFSheet sourceSheet) {
    Objects.requireNonNull(sourceSheet, "sourceSheet must not be null");
    XSSFDrawing drawing = sourceSheet.getDrawingPatriarch();
    if (drawing == null) {
      return new PictureCopySnapshot(List.of());
    }

    List<PictureCopyPlan> pictures = new ArrayList<>();
    for (var shape : drawing.getShapes()) {
      if (shape instanceof XSSFPicture picture) {
        pictures.add(snapshotPicture(picture));
      }
    }
    return new PictureCopySnapshot(pictures);
  }

  void repairCopiedPictures(XSSFSheet targetSheet, PictureCopySnapshot snapshot) {
    Objects.requireNonNull(targetSheet, "targetSheet must not be null");
    Objects.requireNonNull(snapshot, "snapshot must not be null");
    if (snapshot.pictures().isEmpty()) {
      return;
    }

    XSSFDrawing drawing = targetSheet.getDrawingPatriarch();
    if (drawing == null) {
      throw new IllegalStateException(
          "Copied sheet '" + targetSheet.getSheetName() + "' is missing its drawing patriarch");
    }
    List<XSSFPicture> targetPictures =
        drawing.getShapes().stream()
            .filter(XSSFPicture.class::isInstance)
            .map(XSSFPicture.class::cast)
            .toList();
    if (targetPictures.size() != snapshot.pictures().size()) {
      throw new IllegalStateException(
          "Copied sheet '"
              + targetSheet.getSheetName()
              + "' changed its picture count during clone repair");
    }

    for (int index = 0; index < targetPictures.size(); index++) {
      repairCopiedPicture(targetPictures.get(index), snapshot.pictures().get(index));
    }
  }

  private static PictureCopyPlan snapshotPicture(XSSFPicture picture) {
    String objectName = DrawingAdapter.resolvedName(picture);
    DrawingAdapter.PictureReadback readback = DrawingAdapter.requiredPictureReadback(picture);
    return new PictureCopyPlan(
        objectName, readback.picturePart().getPartName().getName(), readback.format());
  }

  private static void repairCopiedPicture(
      XSSFPicture targetPicture, PictureCopyPlan sourcePicture) {
    PackagePart imagePart =
        requiredImagePart(targetPicture.getDrawing().getPackagePart(), sourcePicture);
    String preferredRelationId =
        DrawingAdapter.reusableRelationId(
                targetPicture.getDrawing(),
                DrawingAdapter.pictureRelationId(targetPicture).orElse(null),
                imagePart.getPartName())
            .orElse(null);
    RelationPart relation =
        targetPicture
            .getDrawing()
            .addRelation(
                preferredRelationId,
                ImageCatalog.pictureRelation(sourcePicture.format()),
                ImageCatalog.pictureDataForPart(imagePart));
    DrawingAdapter.setPictureRelationId(targetPicture, relation.getRelationship().getId());
  }

  private static PackagePart requiredImagePart(
      org.apache.poi.openxml4j.opc.PackagePart drawingPart, PictureCopyPlan sourcePicture) {
    try {
      PackagePartName partName = PackagingURIHelper.createPartName(sourcePicture.sourcePartName());
      PackagePart imagePart = drawingPart.getPackage().getPart(partName);
      if (imagePart == null) {
        throw new IllegalStateException(
            "Copied picture '"
                + sourcePicture.objectName()
                + "' is missing image part '"
                + sourcePicture.sourcePartName()
                + "'");
      }
      return imagePart;
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException(
          "Failed to resolve copied picture image part '" + sourcePicture.sourcePartName() + "'",
          exception);
    }
  }

  record PictureCopySnapshot(List<PictureCopyPlan> pictures) {
    PictureCopySnapshot {
      pictures = List.copyOf(pictures);
    }
  }

  record PictureCopyPlan(
      String objectName,
      String sourcePartName,
      dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat format) {
    PictureCopyPlan {
      Objects.requireNonNull(objectName, "objectName must not be null");
      if (objectName.isBlank()) {
        throw new IllegalArgumentException("objectName must not be blank");
      }
      Objects.requireNonNull(sourcePartName, "sourcePartName must not be null");
      if (sourcePartName.isBlank()) {
        throw new IllegalArgumentException("sourcePartName must not be blank");
      }
      Objects.requireNonNull(format, "format must not be null");
    }
  }

  EmbeddedObjectCopySnapshot snapshot(XssfSheet sourceSheet) {
    Objects.requireNonNull(sourceSheet, "sourceSheet must not be null");
    XSSFDrawing drawing = sourceSheet.xssfSheet().getDrawingPatriarch();
    if (drawing == null) {
      return new EmbeddedObjectCopySnapshot(List.of());
    }

    List<EmbeddedObjectCopyPlan> embeddedObjects = new ArrayList<>();
    for (XSSFShape shape : drawing.getShapes()) {
      if (shape instanceof XSSFObjectData objectData) {
        embeddedObjects.add(snapshotEmbeddedObject(sourceSheet.xssfSheet(), objectData));
      }
    }
    return new EmbeddedObjectCopySnapshot(embeddedObjects);
  }

  void repairCopiedEmbeddedObjects(XssfSheet targetSheet, EmbeddedObjectCopySnapshot snapshot) {
    Objects.requireNonNull(targetSheet, "targetSheet must not be null");
    Objects.requireNonNull(snapshot, "snapshot must not be null");
    for (EmbeddedObjectCopyPlan embeddedObject : snapshot.embeddedObjects) {
      repairCopiedEmbeddedObject(targetSheet.xssfSheet(), embeddedObject);
    }
  }

  private static EmbeddedObjectCopyPlan snapshotEmbeddedObject(
      XSSFSheet sourceSheet, XSSFObjectData objectData) {
    String objectName = DrawingAdapter.resolvedName(objectData);
    String packageRelationId = objectData.getOleObject().getId();
    InternalRelationSnapshot packagePart =
        requiredInternalRelation(
            sourceSheet.getPackagePart(), packageRelationId, objectName, "embedded object package");

    String previewSheetRelationId =
        DrawingAdapter.previewSheetRelationId(objectData.getOleObject()).orElse(null);
    @Nullable InternalRelationSnapshot previewSheetPart =
        previewSheetRelationId == null
            ? null
            : requiredInternalRelation(
                sourceSheet.getPackagePart(),
                previewSheetRelationId,
                objectName,
                "embedded object sheet preview");
    String previewDrawingRelationId =
        DrawingAdapter.previewDrawingRelationId(objectData).orElse(null);
    @Nullable InternalRelationSnapshot previewDrawingPart =
        previewDrawingRelationId == null
            ? null
            : requiredInternalRelation(
                objectData.getDrawing().getPackagePart(),
                previewDrawingRelationId,
                objectName,
                "embedded object drawing preview");
    return new EmbeddedObjectCopyPlan(
        objectName, packagePart, previewSheetPart, previewDrawingPart);
  }

  static InternalRelationSnapshot requiredInternalRelation(
      PackagePart sourcePart,
      @Nullable String relationshipId,
      String objectName,
      String relationDescription) {
    if (relationshipId == null || relationshipId.isBlank()) {
      throw new IllegalStateException(
          "Embedded object '" + objectName + "' is missing its " + relationDescription + " id");
    }
    PackageRelationship relationship = sourcePart.getRelationship(relationshipId);
    if (relationship == null || relationship.getTargetMode() == TargetMode.EXTERNAL) {
      throw new IllegalStateException(
          "Embedded object '"
              + objectName
              + "' is missing its "
              + relationDescription
              + " relationship");
    }
    PackagePart part = DrawingAdapter.relatedInternalPart(sourcePart, relationshipId).orElse(null);
    if (part == null) {
      throw new IllegalStateException(
          "Embedded object '" + objectName + "' is missing its " + relationDescription + " part");
    }
    return new InternalRelationSnapshot(
        relationship.getRelationshipType(),
        part.getContentType(),
        part.getPartName().getName(),
        BinaryContent.readback(DrawingAdapter.partBytes(part)));
  }

  private static void repairCopiedEmbeddedObject(
      XSSFSheet targetSheet, EmbeddedObjectCopyPlan embeddedObject) {
    XSSFObjectData targetObjectData =
        requiredEmbeddedObject(targetSheet, embeddedObject.objectName());
    repairPackageRelation(
        targetSheet, targetObjectData, embeddedObject.packagePart(), embeddedObject.objectName());

    String targetPreviewRelationId =
        DrawingAdapter.previewSheetRelationId(targetObjectData.getOleObject()).orElse(null);
    if (embeddedObject.previewSheetPart() != null) {
      repairSheetPreviewRelation(
          targetSheet,
          targetObjectData,
          targetPreviewRelationId,
          embeddedObject.previewSheetPart(),
          embeddedObject.objectName());
    }
    String targetDrawingPreviewRelationId =
        DrawingAdapter.previewDrawingRelationId(targetObjectData).orElse(null);
    if (targetDrawingPreviewRelationId != null && embeddedObject.previewDrawingPart() != null) {
      ensureInternalRelation(
          targetObjectData.getDrawing().getPackagePart(),
          targetDrawingPreviewRelationId,
          embeddedObject.previewDrawingPart(),
          "embedded object drawing preview",
          embeddedObject.objectName());
    }
    repairSheetDrawingRelation(targetSheet);
  }

  private static void repairPackageRelation(
      XSSFSheet targetSheet,
      XSSFObjectData targetObjectData,
      InternalRelationSnapshot sourcePart,
      String objectName) {
    String currentRelationId =
        DrawingAdapter.blankAsOptional(targetObjectData.getOleObject().getId()).orElse(null);
    String repairedRelationId =
        repairWorksheetBoundRelation(
            targetSheet,
            targetObjectData,
            WorksheetRelationRole.OLE_OBJECT,
            currentRelationId,
            sourcePart,
            "embedded object package",
            objectName);
    targetObjectData.getOleObject().setId(repairedRelationId);
  }

  private static void repairSheetPreviewRelation(
      XSSFSheet targetSheet,
      XSSFObjectData targetObjectData,
      @Nullable String currentRelationId,
      InternalRelationSnapshot sourcePart,
      String objectName) {
    String repairedRelationId =
        repairWorksheetBoundRelation(
            targetSheet,
            targetObjectData,
            WorksheetRelationRole.PREVIEW_SHEET,
            currentRelationId,
            sourcePart,
            "embedded object sheet preview",
            objectName);
    DrawingAdapter.setPreviewSheetRelationId(targetObjectData.getOleObject(), repairedRelationId);
  }

  static XSSFObjectData requiredEmbeddedObject(XSSFSheet targetSheet, String objectName) {
    XSSFDrawing drawing = targetSheet.getDrawingPatriarch();
    if (drawing == null) {
      throw new IllegalStateException(
          "Copied sheet '" + targetSheet.getSheetName() + "' is missing its drawing patriarch");
    }
    return drawing.getShapes().stream()
        .filter(XSSFObjectData.class::isInstance)
        .map(XSSFObjectData.class::cast)
        .filter(shape -> objectName.equals(DrawingAdapter.resolvedName(shape)))
        .findFirst()
        .orElseThrow(
            () ->
                new IllegalStateException(
                    "Copied sheet '"
                        + targetSheet.getSheetName()
                        + "' is missing embedded object '"
                        + objectName
                        + "'"));
  }

  static void ensureInternalRelation(
      PackagePart sheetPart,
      String relationshipId,
      InternalRelationSnapshot sourcePart,
      String relationDescription,
      String objectName) {
    @Nullable PackageRelationship existingRelationship = sheetPart.getRelationship(relationshipId);
    @Nullable PackagePart existingPart =
        DrawingAdapter.relatedInternalPart(sheetPart, relationshipId).orElse(null);
    if (matchingRelation(existingRelationship, existingPart, sourcePart)) {
      return;
    }
    if (existingRelationship != null) {
      sheetPart.removeRelationship(relationshipId);
    }
    PackagePart copiedPart =
        createCopiedPart(sheetPart.getPackage(), sourcePart, relationDescription, objectName);
    sheetPart.addRelationship(
        copiedPart.getPartName(),
        TargetMode.INTERNAL,
        sourcePart.relationshipType(),
        relationshipId);
  }

  private static boolean matchingRelation(
      @Nullable PackageRelationship existingRelationship,
      @Nullable PackagePart existingPart,
      InternalRelationSnapshot sourcePart) {
    if (existingRelationship == null || existingPart == null) {
      return false;
    }
    return sourcePart.relationshipType().equals(existingRelationship.getRelationshipType())
        && sourcePart.contentType().equals(existingPart.getContentType())
        && Arrays.equals(sourcePart.bytes().bytes(), DrawingAdapter.partBytes(existingPart));
  }

  static String repairWorksheetBoundRelation(
      XSSFSheet targetSheet,
      XSSFObjectData targetObjectData,
      WorksheetRelationRole relationRole,
      @Nullable String preferredRelationId,
      InternalRelationSnapshot sourcePart,
      String relationDescription,
      String objectName) {
    PackagePart sheetPart = targetSheet.getPackagePart();
    @Nullable PackageRelationship existingRelationship =
        preferredRelationId == null ? null : sheetPart.getRelationship(preferredRelationId);
    @Nullable PackagePart existingPart =
        preferredRelationId == null
            ? null
            : DrawingAdapter.relatedInternalPart(sheetPart, preferredRelationId).orElse(null);
    boolean referencedElsewhere =
        worksheetRelationIdReferencedElsewhere(
            targetSheet, targetObjectData, relationRole, preferredRelationId);
    if (matchingRelation(existingRelationship, existingPart, sourcePart) && !referencedElsewhere) {
      return Objects.requireNonNull(
          preferredRelationId, "preferredRelationId must not be null when relation matches");
    }
    PackagePart copiedPart =
        createCopiedPart(sheetPart.getPackage(), sourcePart, relationDescription, objectName);
    boolean reusePreferredId =
        preferredRelationId != null
            && !referencedElsewhere
            && (existingRelationship == null
                || replaceableWorksheetBoundRelation(
                    existingRelationship, existingPart, sourcePart));
    if (reusePreferredId) {
      String resolvedPreferredRelationId =
          Objects.requireNonNull(preferredRelationId, "preferredRelationId must not be null");
      if (existingRelationship != null) {
        removeExistingRelationAndCleanup(sheetPart, resolvedPreferredRelationId, existingPart);
      }
      return sheetPart
          .addRelationship(
              copiedPart.getPartName(),
              TargetMode.INTERNAL,
              sourcePart.relationshipType(),
              resolvedPreferredRelationId)
          .getId();
    }
    return sheetPart
        .addRelationship(
            copiedPart.getPartName(),
            TargetMode.INTERNAL,
            sourcePart.relationshipType(),
            nextWorksheetRelationId(targetSheet))
        .getId();
  }

  private static boolean replaceableWorksheetBoundRelation(
      PackageRelationship existingRelationship,
      @Nullable PackagePart existingPart,
      InternalRelationSnapshot sourcePart) {
    return sourcePart.relationshipType().equals(existingRelationship.getRelationshipType())
        && (existingPart == null || sourcePart.contentType().equals(existingPart.getContentType()));
  }

  private static void removeExistingRelationAndCleanup(
      PackagePart sheetPart, String relationshipId, @Nullable PackagePart existingPart) {
    sheetPart.removeRelationship(relationshipId);
    if (existingPart != null) {
      PackageRelations.cleanupPackagePartIfUnused(
          sheetPart.getPackage(), existingPart.getPartName());
    }
  }

  static boolean worksheetRelationIdReferencedElsewhere(
      XSSFSheet targetSheet,
      XSSFObjectData targetObjectData,
      WorksheetRelationRole relationRole,
      @Nullable String relationId) {
    String normalizedRelationId = DrawingAdapter.blankAsOptional(relationId).orElse(null);
    if (normalizedRelationId == null) {
      return false;
    }
    var worksheet = targetSheet.getCTWorksheet();
    return worksheetStructureReferencesId(worksheet, normalizedRelationId)
        || (worksheet.isSetOleObjects()
            && oleObjectReferencesIdElsewhere(
                worksheet.getOleObjects().getOleObjectList(),
                targetObjectData.getOleObject(),
                relationRole,
                normalizedRelationId));
  }

  static boolean worksheetStructureReferencesId(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet worksheet,
      String relationId) {
    return (worksheet.isSetDrawing() && relationId.equals(worksheet.getDrawing().getId()))
        || (worksheet.isSetLegacyDrawing()
            && relationId.equals(worksheet.getLegacyDrawing().getId()))
        || (worksheet.isSetLegacyDrawingHF()
            && relationId.equals(worksheet.getLegacyDrawingHF().getId()))
        || (worksheet.isSetDrawingHF() && relationId.equals(worksheet.getDrawingHF().getId()));
  }

  private static boolean oleObjectReferencesIdElsewhere(
      List<org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject> oleObjects,
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject targetOleObject,
      WorksheetRelationRole relationRole,
      String relationId) {
    for (var oleObject : oleObjects) {
      if (referencesOleObjectRelationId(oleObject, targetOleObject, relationRole, relationId)
          || referencesPreviewRelationId(oleObject, targetOleObject, relationRole, relationId)) {
        return true;
      }
    }
    return false;
  }

  private static boolean referencesOleObjectRelationId(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject oleObject,
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject targetOleObject,
      WorksheetRelationRole relationRole,
      String relationId) {
    boolean differentTarget =
        relationRole != WorksheetRelationRole.OLE_OBJECT
            || !isTargetOleObject(oleObject, targetOleObject);
    return differentTarget
        && relationId.equals(DrawingAdapter.blankAsOptional(oleObject.getId()).orElse(null));
  }

  private static boolean referencesPreviewRelationId(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject oleObject,
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject targetOleObject,
      WorksheetRelationRole relationRole,
      String relationId) {
    boolean differentTarget =
        relationRole != WorksheetRelationRole.PREVIEW_SHEET
            || !isTargetOleObject(oleObject, targetOleObject);
    return differentTarget
        && relationId.equals(DrawingAdapter.previewSheetRelationId(oleObject).orElse(null));
  }

  private static boolean isTargetOleObject(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject oleObject,
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject targetOleObject) {
    return oleObject.getShapeId() == targetOleObject.getShapeId();
  }

  private static String nextWorksheetRelationId(XSSFSheet targetSheet) {
    return nextWorksheetRelationId(targetSheet, targetSheet.getPackagePart()::getRelationships);
  }

  static String nextWorksheetRelationId(
      XSSFSheet targetSheet, WorksheetRelationshipSupplier relationshipSupplier) {
    Set<String> reservedIds = referencedWorksheetRelationIds(targetSheet);
    try {
      for (PackageRelationship relationship : relationshipSupplier.relationships()) {
        reservedIds.add(relationship.getId());
      }
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException(
          "Failed to inspect worksheet relationships for copied sheet '"
              + targetSheet.getSheetName()
              + "'",
          exception);
    }
    for (int attempt = 1; ; attempt++) {
      String candidate = "rId" + attempt;
      if (!reservedIds.contains(candidate)) {
        return candidate;
      }
    }
  }

  static Set<String> referencedWorksheetRelationIds(XSSFSheet targetSheet) {
    Set<String> referencedIds = new LinkedHashSet<>();
    var worksheet = targetSheet.getCTWorksheet();
    if (worksheet.isSetDrawing()) {
      addNonBlank(referencedIds, worksheet.getDrawing().getId());
    }
    if (worksheet.isSetLegacyDrawing()) {
      addNonBlank(referencedIds, worksheet.getLegacyDrawing().getId());
    }
    if (worksheet.isSetLegacyDrawingHF()) {
      addNonBlank(referencedIds, worksheet.getLegacyDrawingHF().getId());
    }
    if (worksheet.isSetDrawingHF()) {
      addNonBlank(referencedIds, worksheet.getDrawingHF().getId());
    }
    if (worksheet.isSetOleObjects()) {
      for (var oleObject : worksheet.getOleObjects().getOleObjectList()) {
        addNonBlank(referencedIds, oleObject.getId());
        addNonBlank(referencedIds, DrawingAdapter.previewSheetRelationId(oleObject).orElse(null));
      }
    }
    return referencedIds;
  }

  private static void addNonBlank(Set<String> ids, @Nullable String value) {
    String normalized = DrawingAdapter.blankAsOptional(value).orElse(null);
    if (normalized != null) {
      ids.add(normalized);
    }
  }

  static void repairSheetDrawingRelation(XSSFSheet targetSheet) {
    if (!targetSheet.getCTWorksheet().isSetDrawing()) {
      return;
    }
    String drawingRelationId =
        DrawingAdapter.blankAsOptional(targetSheet.getCTWorksheet().getDrawing().getId())
            .orElse(null);
    if (drawingRelationId == null) {
      return;
    }
    XSSFDrawing drawing = targetSheet.getDrawingPatriarch();
    if (drawing == null) {
      throw new IllegalStateException(
          "Copied sheet '" + targetSheet.getSheetName() + "' is missing its drawing patriarch");
    }
    repairSheetDrawingRelation(targetSheet, drawing);
  }

  static void repairSheetDrawingRelation(XSSFSheet targetSheet, XSSFDrawing drawing) {
    String drawingRelationId =
        DrawingAdapter.blankAsOptional(targetSheet.getCTWorksheet().getDrawing().getId())
            .orElse(null);
    if (drawingRelationId == null) {
      return;
    }
    PackagePart sheetPart = targetSheet.getPackagePart();
    PackageRelationship existingRelationship = sheetPart.getRelationship(drawingRelationId);
    PackagePart existingPart =
        DrawingAdapter.relatedInternalPart(sheetPart, drawingRelationId).orElse(null);
    if (existingRelationship != null
        && existingPart != null
        && org.apache.poi.xssf.usermodel.XSSFRelation.DRAWINGS
            .getRelation()
            .equals(existingRelationship.getRelationshipType())
        && drawing.getPackagePart().getPartName().equals(existingPart.getPartName())) {
      return;
    }
    if (existingRelationship != null) {
      sheetPart.removeRelationship(drawingRelationId);
    }
    sheetPart.addRelationship(
        drawing.getPackagePart().getPartName(),
        TargetMode.INTERNAL,
        org.apache.poi.xssf.usermodel.XSSFRelation.DRAWINGS.getRelation(),
        drawingRelationId);
  }

  static PackagePart createCopiedPart(
      OPCPackage pkg,
      InternalRelationSnapshot sourcePart,
      String relationDescription,
      String objectName) {
    try {
      PackagePartName targetPartName = nextCopiedPartName(pkg, sourcePart.sourcePartName());
      PackagePart copiedPart = pkg.createPart(targetPartName, sourcePart.contentType());
      try (var outputStream = copiedPart.getOutputStream()) {
        outputStream.write(sourcePart.bytes().bytes());
      }
      return copiedPart;
    } catch (IOException | InvalidFormatException exception) {
      throw new IllegalStateException(
          "Failed to copy " + relationDescription + " for embedded object '" + objectName + "'",
          exception);
    }
  }

  static PackagePartName nextCopiedPartName(OPCPackage pkg, String sourcePartName)
      throws InvalidFormatException {
    int extensionIndex = sourcePartName.lastIndexOf('.');
    String baseName =
        extensionIndex >= 0 ? sourcePartName.substring(0, extensionIndex) : sourcePartName;
    String extension = extensionIndex >= 0 ? sourcePartName.substring(extensionIndex) : "";
    for (int attempt = 1; ; attempt++) {
      PackagePartName candidate =
          PackagingURIHelper.createPartName(baseName + "-gridgrind-copy-" + attempt + extension);
      if (!pkg.containPart(candidate)) {
        return candidate;
      }
    }
  }

  /** Opaque embedded-object copy plan carried between sheet snapshot and repair phases. */
  static final class EmbeddedObjectCopySnapshot {
    private final List<EmbeddedObjectCopyPlan> embeddedObjects;

    private EmbeddedObjectCopySnapshot(List<EmbeddedObjectCopyPlan> embeddedObjects) {
      this.embeddedObjects = List.copyOf(embeddedObjects);
    }

    boolean isEmpty() {
      return embeddedObjects.isEmpty();
    }

    int embeddedObjectCount() {
      return embeddedObjects.size();
    }
  }

  record EmbeddedObjectCopyPlan(
      String objectName,
      InternalRelationSnapshot packagePart,
      @Nullable InternalRelationSnapshot previewSheetPart,
      @Nullable InternalRelationSnapshot previewDrawingPart) {
    EmbeddedObjectCopyPlan {
      Objects.requireNonNull(objectName, "objectName must not be null");
      Objects.requireNonNull(packagePart, "packagePart must not be null");
    }
  }

  record InternalRelationSnapshot(
      String relationshipType, String contentType, String sourcePartName, BinaryContent bytes) {
    InternalRelationSnapshot {
      relationshipType = requireNonBlank(relationshipType, "relationshipType");
      contentType = requireNonBlank(contentType, "contentType");
      sourcePartName = requireNonBlank(sourcePartName, "sourcePartName");
      Objects.requireNonNull(bytes, "bytes must not be null");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  /**
   * Identifies which worksheet XML attribute currently owns one copied embedded-object relation.
   */
  enum WorksheetRelationRole {
    OLE_OBJECT,
    PREVIEW_SHEET
  }

  /** Supplies worksheet relationships for id allocation and failure-path coverage. */
  @FunctionalInterface
  interface WorksheetRelationshipSupplier {
    /** Returns the worksheet relationships that currently reserve relation ids. */
    Iterable<PackageRelationship> relationships() throws InvalidFormatException;
  }

  RewrittenChartFormulas rewriteDefinedNameFormulas(XSSFSheet sourceSheet) {
    Objects.requireNonNull(sourceSheet, "sourceSheet must not be null");
    XSSFDrawing drawing = sourceSheet.getDrawingPatriarch();
    if (drawing == null) {
      return new RewrittenChartFormulas(List.of());
    }

    List<FormulaRewrite> rewrites = new ArrayList<>();
    for (XSSFChart chart : drawing.getCharts()) {
      collectFormulaRewrites(sourceSheet, chart, rewrites);
    }
    return new RewrittenChartFormulas(rewrites);
  }

  private static void collectFormulaRewrites(
      XSSFSheet contextSheet, XSSFChart chart, List<FormulaRewrite> rewrites) {
    collectFormulaRewrites(contextSheet, chart.getCTChart().getDomNode(), rewrites);
  }

  private static void collectFormulaRewrites(
      XSSFSheet contextSheet, Node node, List<FormulaRewrite> rewrites) {
    if (isFormulaNode(node)) {
      maybeRewriteDefinedNameFormula(contextSheet, node, rewrites);
    }
    for (Node child = node.getFirstChild(); child != null; child = child.getNextSibling()) {
      collectFormulaRewrites(contextSheet, child, rewrites);
    }
  }

  private static void maybeRewriteDefinedNameFormula(
      XSSFSheet contextSheet, Node formulaNode, List<FormulaRewrite> rewrites) {
    Optional<String> originalFormula = formulaText(formulaNode);
    if (originalFormula.isEmpty() || originalFormula.orElseThrow().isBlank()) {
      return;
    }
    String normalizedFormula = ChartAdapter.normalizeFormula(originalFormula.orElseThrow());
    if (ChartAdapter.resolveDefinedNameReference(contextSheet, normalizedFormula).isEmpty()) {
      return;
    }

    String rewrittenFormula = explicitAreaFormula(contextSheet, normalizedFormula);
    if (originalFormula.orElseThrow().startsWith("=")) {
      rewrittenFormula = "=" + rewrittenFormula;
    }

    setFormulaText(formulaNode, rewrittenFormula);
    rewrites.add(new FormulaRewrite(formulaNode, originalFormula.orElseThrow()));
  }

  private static boolean isFormulaNode(Node node) {
    return node.getNodeType() == Node.ELEMENT_NODE && "f".equals(node.getLocalName());
  }

  private static Optional<String> formulaText(Node formulaNode) {
    return formulaValueNode(formulaNode).map(Node::getNodeValue);
  }

  private static void setFormulaText(Node formulaNode, String formula) {
    formulaValueNode(formulaNode)
        .orElseThrow(
            () -> new IllegalStateException("Chart formula node is missing its text payload"))
        .setNodeValue(formula);
  }

  static Optional<Node> formulaValueNode(Node formulaNode) {
    for (Node child = formulaNode.getFirstChild(); child != null; child = child.getNextSibling()) {
      if (child.getNodeType() == Node.TEXT_NODE || child.getNodeType() == Node.CDATA_SECTION_NODE) {
        return Optional.of(child);
      }
    }
    return Optional.empty();
  }

  static String explicitAreaFormula(XSSFSheet contextSheet, String formula) {
    Objects.requireNonNull(contextSheet, "contextSheet must not be null");
    String normalizedFormula = ChartAdapter.normalizeFormula(formula);
    ResolvedAreaReference resolved =
        ChartAdapter.resolveAreaReference(contextSheet, normalizedFormula);
    return formatResolvedAreaReference(resolved.sheet(), resolved.areaReference());
  }

  private static String formatResolvedAreaReference(XSSFSheet sheet, AreaReference areaReference) {
    AreaReference explicitReference =
        new AreaReference(
            absoluteCellReference(sheet, areaReference.getFirstCell()),
            absoluteCellReference(sheet, areaReference.getLastCell()),
            SpreadsheetVersion.EXCEL2007);
    return explicitReference.formatAsString();
  }

  private static CellReference absoluteCellReference(XSSFSheet sheet, CellReference reference) {
    return new CellReference(
        sheet.getSheetName(), reference.getRow(), reference.getCol(), true, true);
  }

  /** Opaque batch of source-formula rewrites applied only for the duration of one sheet clone. */
  static final class RewrittenChartFormulas {
    private final List<FormulaRewrite> rewrites;

    private RewrittenChartFormulas(List<FormulaRewrite> rewrites) {
      this.rewrites = List.copyOf(rewrites);
    }

    boolean isEmpty() {
      return rewrites.isEmpty();
    }

    int rewriteCount() {
      return rewrites.size();
    }

    void restoreSourceFormulas() {
      for (FormulaRewrite rewrite : rewrites) {
        rewrite.restore();
      }
    }
  }

  record FormulaRewrite(Node formulaNode, String originalFormula) {
    FormulaRewrite {
      Objects.requireNonNull(formulaNode, "formulaNode must not be null");
      Objects.requireNonNull(originalFormula, "originalFormula must not be null");
    }

    void restore() {
      setFormulaText(formulaNode, originalFormula);
    }
  }

  /** Ensures a source sheet is internally consistent before handing it to POI clone logic. */
  void prepareSourceSheetForClone(XSSFSheet sourceSheet) {
    Objects.requireNonNull(sourceSheet, "sourceSheet must not be null");
    materializeExternalHyperlinkRelationships(sourceSheet);
  }

  private static void materializeExternalHyperlinkRelationships(XSSFSheet sourceSheet) {
    for (XSSFHyperlink hyperlink : List.copyOf(sourceSheet.getHyperlinkList())) {
      materializeExternalHyperlinkRelationship(sourceSheet, hyperlink)
          .ifPresent(
              materialized -> {
                sourceSheet.removeHyperlink(hyperlink);
                sourceSheet.addHyperlink(materialized);
              });
    }
  }

  private static Optional<XSSFHyperlink> materializeExternalHyperlinkRelationship(
      XSSFSheet sourceSheet, XSSFHyperlink hyperlink) {
    if (!hyperlink.needsRelationToo()) {
      return Optional.empty();
    }
    String relationshipId = hyperlink.getCTHyperlink().getId();
    if (relationshipId != null
        && !relationshipId.isBlank()
        && sourceSheet.getPackagePart().getRelationship(relationshipId) != null) {
      return Optional.empty();
    }
    return Optional.of(
        MaterializedHyperlink.wrapWithMaterializedRelation(
            hyperlink, sourceSheet.getPackagePart()));
  }

  /**
   * POI cloneSheet consumes relation-backed XSSFHyperlink objects rather than raw CT hyperlink ids.
   * Rehydrating through the public copy constructor and protected relation materializer keeps the
   * hyperlink list and package relationships aligned without private access.
   */
  private static final class MaterializedHyperlink extends XSSFHyperlink {
    private MaterializedHyperlink(CTHyperlink ctHyperlink, PackageRelationship relationship) {
      super(ctHyperlink, relationship);
    }

    private static XSSFHyperlink wrapWithMaterializedRelation(
        XSSFHyperlink hyperlink, PackagePart sheetPart) {
      String relationshipId = hyperlink.getCTHyperlink().getId();
      String requestedRelationshipId =
          relationshipId == null || relationshipId.isBlank() ? null : relationshipId;
      PackageRelationship relationship =
          sheetPart.addExternalRelationship(
              hyperlink.getAddress(),
              XSSFRelation.SHEET_HYPERLINKS.getRelation(),
              requestedRelationshipId);
      hyperlink.getCTHyperlink().setId(relationship.getId());
      return new MaterializedHyperlink(hyperlink.getCTHyperlink(), relationship);
    }
  }
}
