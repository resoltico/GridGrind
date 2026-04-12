package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

/**
 * Resolves defined-name formulas into typed GridGrind range targets when normalization is possible.
 */
public final class ExcelNamedRangeTargets {
  private ExcelNamedRangeTargets() {}

  /**
   * Returns the typed target for a defined-name formula when it resolves to one contiguous cell or
   * range.
   */
  public static Optional<ExcelNamedRangeTarget> resolveTarget(
      String refersToFormula, ExcelNamedRangeScope scope) {
    Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
    Objects.requireNonNull(scope, "scope must not be null");

    if (refersToFormula.isBlank()) {
      return Optional.empty();
    }
    String normalizedFormula = normalizeAreaFormulaForPoi(refersToFormula);
    if (!AreaReference.isContiguous(normalizedFormula)) {
      return Optional.empty();
    }
    try {
      AreaReference areaReference =
          new AreaReference(normalizedFormula, SpreadsheetVersion.EXCEL2007);
      String sheetName = resolveSheetName(areaReference, scope);
      if (sheetName == null) {
        return Optional.empty();
      }
      CellReference firstCell = areaReference.getFirstCell();
      CellReference lastCell = areaReference.getLastCell();
      String range =
          firstCell.getRow() == lastCell.getRow() && firstCell.getCol() == lastCell.getCol()
              ? new CellReference(firstCell.getRow(), firstCell.getCol()).formatAsString()
              : new CellReference(firstCell.getRow(), firstCell.getCol()).formatAsString()
                  + ":"
                  + new CellReference(lastCell.getRow(), lastCell.getCol()).formatAsString();
      return Optional.of(new ExcelNamedRangeTarget(sheetName, range));
    } catch (RuntimeException exception) {
      return Optional.empty();
    }
  }

  static String normalizeAreaFormulaForPoi(String refersToFormula) {
    int separatorIndex = refersToFormula.indexOf(':');
    if (separatorIndex < 0) {
      return refersToFormula;
    }
    String firstReference = refersToFormula.substring(0, separatorIndex);
    String secondReference = refersToFormula.substring(separatorIndex + 1);
    int firstBangIndex = firstReference.lastIndexOf('!');
    int secondBangIndex = secondReference.lastIndexOf('!');
    if (firstBangIndex < 0 || secondBangIndex < 0) {
      return refersToFormula;
    }
    String firstSheetPrefix = firstReference.substring(0, firstBangIndex + 1);
    String secondSheetPrefix = secondReference.substring(0, secondBangIndex + 1);
    if (!firstSheetPrefix.equals(secondSheetPrefix)) {
      return refersToFormula;
    }
    return firstReference + ":" + secondReference.substring(secondBangIndex + 1);
  }

  private static String resolveSheetName(AreaReference areaReference, ExcelNamedRangeScope scope) {
    String sheetName = areaReference.getFirstCell().getSheetName();
    if (sheetName != null) {
      return sheetName;
    }
    return switch (scope) {
      case ExcelNamedRangeScope.WorkbookScope _ -> null;
      case ExcelNamedRangeScope.SheetScope sheetScope -> sheetScope.sheetName();
    };
  }
}
