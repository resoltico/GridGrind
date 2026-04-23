package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;

/** Copies sheets by combining POI XSSF cloning with GridGrind-owned repair passes. */
final class ExcelSheetCopyController {
  private final ExcelAutofilterController autofilterController = new ExcelAutofilterController();
  private final ExcelSheetClonePreparationSupport clonePreparationSupport =
      new ExcelSheetClonePreparationSupport();
  private final ExcelSheetCopyEmbeddedObjectSupport embeddedObjectCopySupport =
      new ExcelSheetCopyEmbeddedObjectSupport();
  private final ExcelSheetCopyPictureSupport pictureCopySupport =
      new ExcelSheetCopyPictureSupport();

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
    ExcelSheet sourceSheet = workbook.sheet(sourceSheetName);
    ExcelSheetCopyPictureSupport.CopySnapshot pictures =
        pictureCopySupport.snapshot(sourcePoiSheet);
    ExcelSheetCopySnapshot snapshot = snapshot(workbook, sourceSheetName, sourcePoiSheet);
    ExcelSheetCopyEmbeddedObjectSupport.CopySnapshot embeddedObjects =
        embeddedObjectCopySupport.snapshot(sourceSheet);
    clonePreparationSupport.prepareSourceSheetForClone(sourcePoiSheet);
    int sourceSheetIndex = workbook.xssfWorkbook().getSheetIndex(sourcePoiSheet);
    XSSFSheet targetPoiSheet = workbook.xssfWorkbook().cloneSheet(sourceSheetIndex, newSheetName);
    ExcelSheet targetSheet = workbook.sheet(newSheetName);
    List<String> transientCopiedTableNames =
        targetPoiSheet.getTables().stream()
            .map(table -> Objects.requireNonNullElse(table.getName(), ""))
            .toList();

    pictureCopySupport.repairCopiedPictures(targetPoiSheet, pictures);
    retargetCopiedSheetFormulas(workbook, sourceSheetName, newSheetName, targetPoiSheet);
    ExcelSheetCopySupport.replaceDataValidations(
        snapshot.dataValidations(), targetPoiSheet, workbook, sourceSheetName, newSheetName);
    ExcelSheetCopySupport.replaceConditionalFormatting(
        snapshot.conditionalFormattingBlocks(),
        targetSheet,
        workbook,
        sourceSheetName,
        newSheetName);
    ExcelSheetCopySupport.repairComments(snapshot.comments(), targetSheet);
    embeddedObjectCopySupport.repairCopiedEmbeddedObjects(targetSheet, embeddedObjects);
    ExcelSheetCopySupport.repairPrintLayout(snapshot.printLayout(), targetSheet);
    ExcelSheetCopySupport.replaceTables(workbook, newSheetName, snapshot.tables());
    ExcelTableStructuredReferenceRetargetSupport.retargetFormulaCells(
        targetPoiSheet,
        transientCopiedTableNames,
        targetPoiSheet.getTables().stream()
            .map(table -> Objects.requireNonNullElse(table.getName(), ""))
            .toList());
    ExcelTableCalculatedColumnCanonicalizer.canonicalizeSheet(targetPoiSheet);
    ExcelSheetCopySupport.replaceAutofilter(snapshot.sheetAutofilter(), targetSheet);
    ExcelSheetCopySupport.replaceLocalNamedRanges(
        workbook, sourceSheetName, newSheetName, snapshot.localNamedRanges());
    ExcelSheetCopySupport.copyProtection(sourcePoiSheet, targetPoiSheet);

    int targetSheetIndex = workbook.xssfWorkbook().getSheetIndex(targetPoiSheet);
    workbook.xssfWorkbook().setSheetVisibility(targetSheetIndex, SheetVisibility.VISIBLE);
    targetPoiSheet.setSelected(false);
    switch (position) {
      case ExcelSheetCopyPosition.AppendAtEnd _ -> {
        // cloneSheet already appends at the end.
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
        ExcelSheetCopySupport.copyableLocalNames(workbook.namedRanges(), sheetName),
        sourceSheet.layout(),
        sourceSheet.printLayout(),
        sourceSheet.mergedRegions(),
        sourceSheet.comments(new ExcelCellSelection.AllUsedCells()),
        ExcelSheetCopySupport.copiedDataValidations(sourcePoiSheet),
        ExcelSheetCopySupport.supportedConditionalFormatting(
            sourceSheet.conditionalFormatting(new ExcelRangeSelection.All()), sheetName),
        ExcelSheetCopySupport.sheetOwnedAutofilter(
            autofilterController.sheetOwnedAutofilters(sourcePoiSheet)),
        ExcelSheetCopySupport.tablesOnSheet(sourcePoiSheet));
  }

  private static void retargetCopiedSheetFormulas(
      ExcelWorkbook workbook, String sourceSheetName, String newSheetName, XSSFSheet targetSheet) {
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
            ExcelFormulaSheetRenameSupport.renameSheet(
                workbook.xssfWorkbook(),
                cell.getCellFormula(),
                FormulaType.CELL,
                targetSheetIndex,
                sourceSheetName,
                newSheetName);
        if (!rewritten.equals(cell.getCellFormula())) {
          ExcelFormulaWriteSupport.setRewrittenFormula(
              cell, rewritten, "Copied-sheet formula retargeting");
        }
      }
    }
  }

  private static boolean mayReferenceCopiedSheet(String formula, String sourceSheetName) {
    String quotedSourceSheetName = "'" + sourceSheetName.replace("'", "''") + "'";
    return formula.contains(sourceSheetName + "!")
        || formula.contains(quotedSourceSheetName + "!")
        || formula.contains(sourceSheetName + ":")
        || formula.contains(quotedSourceSheetName + ":");
  }

  static void deleteLocalNamedRanges(ExcelWorkbook workbook, String sheetName) {
    ExcelSheetCopySupport.deleteLocalNamedRanges(workbook, sheetName);
  }

  /** Returns copyable sheet-local names scoped to the requested source sheet. */
  static List<ExcelNamedRangeSnapshot> copyableLocalNames(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    return ExcelSheetCopySupport.copyableLocalNames(namedRanges, sheetName);
  }

  /** Retained for focused tests covering explicit range-backed local names only. */
  static List<ExcelNamedRangeSnapshot.RangeSnapshot> copyableLocalRangeNames(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    return ExcelSheetCopySupport.copyableLocalRangeNames(namedRanges, sheetName);
  }

  /** Local names owned by a copied sheet are now fully copyable. */
  static void requireNoUncopyableLocalNamedRanges(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    ExcelSheetCopySupport.requireNoUncopyableLocalNamedRanges(namedRanges, sheetName);
  }

  /** Sheet copy now supports tables and no longer rejects them up front. */
  static void requireNoTables(XSSFSheet sheet, String sheetName) {
    ExcelSheetCopySupport.requireNoTables(sheet, sheetName);
  }

  /** Returns only copyable conditional-formatting blocks for one source sheet. */
  static List<ExcelConditionalFormattingBlockDefinition> supportedConditionalFormatting(
      List<ExcelConditionalFormattingBlockSnapshot> blocks, String sourceSheetName) {
    return ExcelSheetCopySupport.supportedConditionalFormatting(blocks, sourceSheetName);
  }

  /** Converts one persisted conditional-formatting rule into a copyable authoring rule. */
  static ExcelConditionalFormattingRule copyableRule(
      ExcelConditionalFormattingRuleSnapshot rule, String sourceSheetName) {
    return ExcelSheetCopySupport.copyableRule(rule, sourceSheetName);
  }

  /** Converts one supported differential style into its copyable authoring form. */
  static ExcelDifferentialStyle copyableStyle(
      ExcelDifferentialStyleSnapshot style, String sourceSheetName) {
    return ExcelSheetCopySupport.copyableStyle(style, sourceSheetName);
  }

  /** Returns the one sheet-owned autofilter snapshot, or null when absent. */
  static ExcelAutofilterSnapshot.SheetOwned sheetOwnedAutofilter(
      List<ExcelAutofilterSnapshot> autofilters) {
    return ExcelSheetCopySupport.sheetOwnedAutofilter(autofilters);
  }

  /** Retained for focused tests around the legacy range-only sheet autofilter path. */
  static Optional<String> sheetOwnedAutofilterRange(List<ExcelAutofilterSnapshot> autofilters) {
    return ExcelSheetCopySupport.sheetOwnedAutofilterRange(autofilters);
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
