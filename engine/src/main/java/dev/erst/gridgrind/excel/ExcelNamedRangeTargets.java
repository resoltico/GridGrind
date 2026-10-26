package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.jspecify.annotations.Nullable;

/**
 * Resolves defined-name formulas into typed GridGrind range targets when normalization is possible.
 */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelNamedRangeTargets {
  private ExcelNamedRangeTargets() {}

  /**
   * Returns the typed target for a defined-name formula when it resolves to one contiguous cell or
   * range.
   */
  public static String toFormula(NamedRangeTarget target) {
    return switch (target) {
      case NamedRangeTarget.Formula formula -> formula.formula();
      case NamedRangeTarget.Range range -> rangeFormula(range.sheetName(), range.range());
    };
  }

  /** Formats a concrete range for a destination sheet at the POI boundary. */
  public static String rangeFormula(
      dev.erst.gridgrind.core.vocabulary.SheetName sheetName,
      dev.erst.gridgrind.core.vocabulary.RangeRef range) {
    ExcelRange parsed = ExcelRange.parse(range.value());
    String first =
        new CellReference(sheetName.value(), parsed.firstRow(), parsed.firstColumn(), true, true)
            .formatAsString();
    return parsed.rowCount() == 1 && parsed.columnCount() == 1
        ? first
        : first
            + ":"
            + new CellReference(parsed.lastRow(), parsed.lastColumn(), true, true).formatAsString();
  }

  public static Optional<NamedRangeTarget> resolveTarget(
      String refersToFormula, NamedRangeScopeReport scope) {
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
      return Optional.of(
          NamedRangeTarget.range(
              new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName),
              new dev.erst.gridgrind.core.vocabulary.RangeRef(range)));
    } catch (RuntimeException exception) {
      return Optional.empty();
    }
  }

  public static String normalizeAreaFormulaForPoi(String refersToFormula) {
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

  private static @Nullable String resolveSheetName(
      AreaReference areaReference, NamedRangeScopeReport scope) {
    String sheetName = areaReference.getFirstCell().getSheetName();
    if (sheetName != null) {
      return sheetName;
    }
    return switch (scope) {
      case NamedRangeScopeReport.Workbook _ -> null;
      case NamedRangeScopeReport.Sheet sheetScope -> sheetScope.sheetName().value();
    };
  }
}
