package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
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
  private final ExcelSheetCloneChartPreparationSupport cloneChartPreparationSupport =
      new ExcelSheetCloneChartPreparationSupport();
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
      SheetCopyPosition position) {
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
    ExcelSheetCloneChartPreparationSupport.RewrittenChartFormulas rewrittenChartFormulas =
        cloneChartPreparationSupport.rewriteDefinedNameFormulas(sourcePoiSheet);
    clonePreparationSupport.prepareSourceSheetForClone(sourcePoiSheet);
    int sourceSheetIndex = workbook.xssfWorkbook().getSheetIndex(sourcePoiSheet);
    XSSFSheet targetPoiSheet;
    try {
      targetPoiSheet = workbook.xssfWorkbook().cloneSheet(sourceSheetIndex, newSheetName);
    } finally {
      rewrittenChartFormulas.restoreSourceFormulas();
    }
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
      case SheetCopyPosition.AppendAtEnd _ -> {
        // cloneSheet already appends at the end.
      }
      case SheetCopyPosition.AtIndex atIndex -> {
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
        ExcelSheetCopySupport.copyableLocalNames(workbook.names().namedRanges(), sheetName),
        sourceSheet.layout().printLayout(),
        sourceSheet.annotations().comments(new ExcelCellSelection.AllUsedCells()),
        ExcelSheetCopySupport.copiedDataValidations(sourcePoiSheet),
        ExcelSheetCopyConditionalFormattingSupport.snapshot(sourcePoiSheet, sheetName),
        ExcelSheetCopySupport.sheetOwnedAutofilter(
            autofilterController.sheetOwnedAutofilters(sourcePoiSheet)),
        ExcelSheetCopySupport.tablesOnSheet(sourcePoiSheet));
  }

  static void retargetCopiedSheetFormulas(
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

  static boolean mayReferenceCopiedSheet(String formula, String sourceSheetName) {
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
  static List<NamedRangeReport> copyableLocalNames(
      List<NamedRangeReport> namedRanges, String sheetName) {
    return ExcelSheetCopySupport.copyableLocalNames(namedRanges, sheetName);
  }

  /** Returns the one sheet-owned autofilter snapshot when present. */
  static Optional<AutofilterEntryReport.SheetOwned> sheetOwnedAutofilter(
      List<AutofilterEntryReport> autofilters) {
    return ExcelSheetCopySupport.sheetOwnedAutofilter(autofilters);
  }

  /** Immutable copy plan for the sheet-local structures GridGrind can reproduce safely. */
  private record ExcelSheetCopySnapshot(
      List<NamedRangeReport> localNamedRanges,
      ResolvedPrintLayout printLayout,
      List<CellCommentReport> comments,
      List<CTDataValidation> dataValidations,
      List<org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting>
          conditionalFormattingBlocks,
      Optional<AutofilterEntryReport.SheetOwned> sheetAutofilter,
      List<TableEntryReport> tables) {
    private ExcelSheetCopySnapshot {
      localNamedRanges = List.copyOf(localNamedRanges);
      Objects.requireNonNull(printLayout, "printLayout must not be null");
      comments = List.copyOf(comments);
      dataValidations = List.copyOf(dataValidations);
      conditionalFormattingBlocks = List.copyOf(conditionalFormattingBlocks);
      Objects.requireNonNull(sheetAutofilter, "sheetAutofilter must not be null");
      tables = List.copyOf(tables);
    }
  }
}
