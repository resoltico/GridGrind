package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Shared helpers that keep {@link ProblemContext} variants focused on public contract shape. */
final class ProblemContextSupport {
  private ProblemContextSupport() {}

  static ProblemContextWorkbookSurfaces.ProblemLocation mergeLocation(
      ProblemContextWorkbookSurfaces.ProblemLocation current,
      ProblemContextWorkbookSurfaces.ProblemLocation discovered) {
    Objects.requireNonNull(current, "current must not be null");
    Objects.requireNonNull(discovered, "discovered must not be null");
    if (current instanceof ProblemContextWorkbookSurfaces.ProblemLocation.Unknown) {
      return discovered;
    }
    if (discovered instanceof ProblemContextWorkbookSurfaces.ProblemLocation.Unknown) {
      return current;
    }
    if (current instanceof ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell) {
      return current;
    }
    if (current instanceof ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange) {
      return current;
    }
    if (current instanceof ProblemContextWorkbookSurfaces.ProblemLocation.NamedRange) {
      return current;
    }
    if (current instanceof ProblemContextWorkbookSurfaces.ProblemLocation.Range) {
      return current;
    }
    if (current instanceof ProblemContextWorkbookSurfaces.ProblemLocation.RangeOnly currentRange) {
      return mergeRangeOnlyLocation(currentRange, discovered);
    }
    if (current instanceof ProblemContextWorkbookSurfaces.ProblemLocation.Address currentAddress) {
      return mergeAddressLocation(currentAddress, discovered);
    }
    if (current instanceof ProblemContextWorkbookSurfaces.ProblemLocation.Sheet currentSheet) {
      return mergeSheetLocation(currentSheet, discovered);
    }
    return mergeCellLocation(
        (ProblemContextWorkbookSurfaces.ProblemLocation.Cell) current, discovered);
  }

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static ProblemContextWorkbookSurfaces.ProblemLocation mergeSheetLocation(
      ProblemContextWorkbookSurfaces.ProblemLocation.Sheet current,
      ProblemContextWorkbookSurfaces.ProblemLocation discovered) {
    if (discovered instanceof ProblemContextWorkbookSurfaces.ProblemLocation.Sheet) {
      return current;
    }
    if (discovered
        instanceof ProblemContextWorkbookSurfaces.ProblemLocation.Address discoveredAddress) {
      return new ProblemContextWorkbookSurfaces.ProblemLocation.Cell(
          current.sheetName(), discoveredAddress.address());
    }
    if (discovered instanceof ProblemContextWorkbookSurfaces.ProblemLocation.Cell discoveredCell) {
      return new ProblemContextWorkbookSurfaces.ProblemLocation.Cell(
          current.sheetName(), discoveredCell.address());
    }
    if (discovered
        instanceof ProblemContextWorkbookSurfaces.ProblemLocation.RangeOnly discoveredRange) {
      return new ProblemContextWorkbookSurfaces.ProblemLocation.Range(
          current.sheetName(), discoveredRange.range());
    }
    if (discovered
        instanceof ProblemContextWorkbookSurfaces.ProblemLocation.Range discoveredRange) {
      return new ProblemContextWorkbookSurfaces.ProblemLocation.Range(
          current.sheetName(), discoveredRange.range());
    }
    if (discovered
        instanceof ProblemContextWorkbookSurfaces.ProblemLocation.NamedRange discoveredNamedRange) {
      return discoveredNamedRange;
    }
    if (discovered
        instanceof
        ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange discoveredNamedRange) {
      return discoveredNamedRange;
    }
    ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell discoveredFormulaCell =
        (ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell) discovered;
    return new ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell(
        current.sheetName(), discoveredFormulaCell.address(), discoveredFormulaCell.formula());
  }

  private static ProblemContextWorkbookSurfaces.ProblemLocation mergeCellLocation(
      ProblemContextWorkbookSurfaces.ProblemLocation.Cell current,
      ProblemContextWorkbookSurfaces.ProblemLocation discovered) {
    if (discovered
        instanceof
        ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell discoveredFormulaCell) {
      return new ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell(
          current.sheetName(), current.address(), discoveredFormulaCell.formula());
    }
    return current;
  }

  private static ProblemContextWorkbookSurfaces.ProblemLocation mergeAddressLocation(
      ProblemContextWorkbookSurfaces.ProblemLocation.Address current,
      ProblemContextWorkbookSurfaces.ProblemLocation discovered) {
    if (discovered
        instanceof
        ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell discoveredFormulaCell) {
      return new ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell(
          discoveredFormulaCell.sheetName(), current.address(), discoveredFormulaCell.formula());
    }
    if (discovered
        instanceof ProblemContextWorkbookSurfaces.ProblemLocation.Sheet discoveredSheet) {
      return new ProblemContextWorkbookSurfaces.ProblemLocation.Cell(
          discoveredSheet.sheetName(), current.address());
    }
    return current;
  }

  private static ProblemContextWorkbookSurfaces.ProblemLocation mergeRangeOnlyLocation(
      ProblemContextWorkbookSurfaces.ProblemLocation.RangeOnly current,
      ProblemContextWorkbookSurfaces.ProblemLocation discovered) {
    if (discovered
        instanceof ProblemContextWorkbookSurfaces.ProblemLocation.Sheet discoveredSheet) {
      return new ProblemContextWorkbookSurfaces.ProblemLocation.Range(
          discoveredSheet.sheetName(), current.range());
    }
    if (discovered instanceof ProblemContextWorkbookSurfaces.ProblemLocation.Range
        || discovered instanceof ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell
        || discovered instanceof ProblemContextWorkbookSurfaces.ProblemLocation.Cell
        || discovered instanceof ProblemContextWorkbookSurfaces.ProblemLocation.NamedRange
        || discovered instanceof ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange) {
      return discovered;
    }
    return current;
  }
}
