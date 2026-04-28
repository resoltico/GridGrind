package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Shared helpers that keep {@link ProblemContext} variants focused on public contract shape. */
final class ProblemContextSupport {
  private ProblemContextSupport() {}

  static ProblemContext.ProblemLocation mergeLocation(
      ProblemContext.ProblemLocation current, ProblemContext.ProblemLocation discovered) {
    Objects.requireNonNull(current, "current must not be null");
    Objects.requireNonNull(discovered, "discovered must not be null");
    if (current instanceof ProblemContext.ProblemLocation.Unknown) {
      return discovered;
    }
    if (discovered instanceof ProblemContext.ProblemLocation.Unknown) {
      return current;
    }
    if (current instanceof ProblemContext.ProblemLocation.FormulaCell) {
      return current;
    }
    if (current instanceof ProblemContext.ProblemLocation.SheetNamedRange) {
      return current;
    }
    if (current instanceof ProblemContext.ProblemLocation.NamedRange) {
      return current;
    }
    if (current instanceof ProblemContext.ProblemLocation.Range) {
      return current;
    }
    if (current instanceof ProblemContext.ProblemLocation.RangeOnly currentRange) {
      return mergeRangeOnlyLocation(currentRange, discovered);
    }
    if (current instanceof ProblemContext.ProblemLocation.Address currentAddress) {
      return mergeAddressLocation(currentAddress, discovered);
    }
    if (current instanceof ProblemContext.ProblemLocation.Sheet currentSheet) {
      return mergeSheetLocation(currentSheet, discovered);
    }
    return mergeCellLocation((ProblemContext.ProblemLocation.Cell) current, discovered);
  }

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static ProblemContext.ProblemLocation mergeSheetLocation(
      ProblemContext.ProblemLocation.Sheet current, ProblemContext.ProblemLocation discovered) {
    if (discovered instanceof ProblemContext.ProblemLocation.Sheet) {
      return current;
    }
    if (discovered instanceof ProblemContext.ProblemLocation.Address discoveredAddress) {
      return new ProblemContext.ProblemLocation.Cell(
          current.sheetName(), discoveredAddress.address());
    }
    if (discovered instanceof ProblemContext.ProblemLocation.Cell discoveredCell) {
      return new ProblemContext.ProblemLocation.Cell(current.sheetName(), discoveredCell.address());
    }
    if (discovered instanceof ProblemContext.ProblemLocation.RangeOnly discoveredRange) {
      return new ProblemContext.ProblemLocation.Range(current.sheetName(), discoveredRange.range());
    }
    if (discovered instanceof ProblemContext.ProblemLocation.Range discoveredRange) {
      return new ProblemContext.ProblemLocation.Range(current.sheetName(), discoveredRange.range());
    }
    if (discovered instanceof ProblemContext.ProblemLocation.NamedRange discoveredNamedRange) {
      return discoveredNamedRange;
    }
    if (discovered instanceof ProblemContext.ProblemLocation.SheetNamedRange discoveredNamedRange) {
      return discoveredNamedRange;
    }
    ProblemContext.ProblemLocation.FormulaCell discoveredFormulaCell =
        (ProblemContext.ProblemLocation.FormulaCell) discovered;
    return new ProblemContext.ProblemLocation.FormulaCell(
        current.sheetName(), discoveredFormulaCell.address(), discoveredFormulaCell.formula());
  }

  private static ProblemContext.ProblemLocation mergeCellLocation(
      ProblemContext.ProblemLocation.Cell current, ProblemContext.ProblemLocation discovered) {
    if (discovered instanceof ProblemContext.ProblemLocation.FormulaCell discoveredFormulaCell) {
      return new ProblemContext.ProblemLocation.FormulaCell(
          current.sheetName(), current.address(), discoveredFormulaCell.formula());
    }
    return current;
  }

  private static ProblemContext.ProblemLocation mergeAddressLocation(
      ProblemContext.ProblemLocation.Address current, ProblemContext.ProblemLocation discovered) {
    if (discovered instanceof ProblemContext.ProblemLocation.FormulaCell discoveredFormulaCell) {
      return new ProblemContext.ProblemLocation.FormulaCell(
          discoveredFormulaCell.sheetName(), current.address(), discoveredFormulaCell.formula());
    }
    if (discovered instanceof ProblemContext.ProblemLocation.Sheet discoveredSheet) {
      return new ProblemContext.ProblemLocation.Cell(
          discoveredSheet.sheetName(), current.address());
    }
    return current;
  }

  private static ProblemContext.ProblemLocation mergeRangeOnlyLocation(
      ProblemContext.ProblemLocation.RangeOnly current, ProblemContext.ProblemLocation discovered) {
    if (discovered instanceof ProblemContext.ProblemLocation.Sheet discoveredSheet) {
      return new ProblemContext.ProblemLocation.Range(discoveredSheet.sheetName(), current.range());
    }
    if (discovered instanceof ProblemContext.ProblemLocation.Range
        || discovered instanceof ProblemContext.ProblemLocation.FormulaCell
        || discovered instanceof ProblemContext.ProblemLocation.Cell
        || discovered instanceof ProblemContext.ProblemLocation.NamedRange
        || discovered instanceof ProblemContext.ProblemLocation.SheetNamedRange) {
      return discovered;
    }
    return current;
  }
}
