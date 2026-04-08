package dev.erst.gridgrind.protocol.exec;

import dev.erst.gridgrind.protocol.dto.CellInput;
import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.dto.RequestWarning;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.stream.IntStream;

/** Collects non-fatal request warnings that can be derived directly from one parsed request. */
final class GridGrindRequestWarnings {
  private GridGrindRequestWarnings() {}

  static List<RequestWarning> collect(GridGrindRequest request) {
    Objects.requireNonNull(request, "request must not be null");

    Set<String> spacedSheetNames = sameRequestSpacedSheetNames(request.operations());
    if (spacedSheetNames.isEmpty()) {
      return List.of();
    }

    return IntStream.range(0, request.operations().size())
        .mapToObj(
            operationIndex ->
                warningFor(
                    request.operations().get(operationIndex), operationIndex, spacedSheetNames))
        .filter(Objects::nonNull)
        .toList();
  }

  private static Set<String> sameRequestSpacedSheetNames(List<WorkbookOperation> operations) {
    Set<String> sheetNames = new LinkedHashSet<>();
    for (WorkbookOperation operation : operations) {
      switch (operation) {
        case WorkbookOperation.EnsureSheet ensureSheet ->
            addIfSpaced(sheetNames, ensureSheet.sheetName());
        case WorkbookOperation.RenameSheet renameSheet ->
            addIfSpaced(sheetNames, renameSheet.newSheetName());
        case WorkbookOperation.CopySheet copySheet ->
            addIfSpaced(sheetNames, copySheet.newSheetName());
        default -> {}
      }
    }
    return sheetNames;
  }

  private static void addIfSpaced(Set<String> sheetNames, String sheetName) {
    if (sheetName.indexOf(' ') >= 0) {
      sheetNames.add(sheetName);
    }
  }

  private static void collectOffendingSheetNames(
      WorkbookOperation operation, Set<String> spacedSheetNames, Set<String> offendingSheetNames) {
    switch (operation) {
      case WorkbookOperation.SetCell setCell ->
          collectFromCellInput(setCell.value(), spacedSheetNames, offendingSheetNames);
      case WorkbookOperation.SetRange setRange -> {
        for (List<CellInput> row : setRange.rows()) {
          for (CellInput cellInput : row) {
            collectFromCellInput(cellInput, spacedSheetNames, offendingSheetNames);
          }
        }
      }
      case WorkbookOperation.AppendRow appendRow -> {
        for (CellInput cellInput : appendRow.values()) {
          collectFromCellInput(cellInput, spacedSheetNames, offendingSheetNames);
        }
      }
      default -> {}
    }
  }

  private static RequestWarning warningFor(
      WorkbookOperation operation, int operationIndex, Set<String> spacedSheetNames) {
    Set<String> offendingSheetNames = new LinkedHashSet<>();
    collectOffendingSheetNames(operation, spacedSheetNames, offendingSheetNames);
    if (offendingSheetNames.isEmpty()) {
      return null;
    }
    return new RequestWarning(
        operationIndex,
        operation.operationType(),
        "Formula references same-request sheet names with spaces without single quotes: "
            + String.join(", ", offendingSheetNames)
            + ". Use 'Sheet Name'!A1 syntax.");
  }

  private static void collectFromCellInput(
      CellInput input, Set<String> spacedSheetNames, Set<String> offendingSheetNames) {
    if (!(input instanceof CellInput.Formula formula)) {
      return;
    }
    String maskedFormula = maskDoubleQuotedStrings(formula.formula());
    for (String sheetName : spacedSheetNames) {
      if (containsUnquotedSheetReference(maskedFormula, sheetName)) {
        offendingSheetNames.add(sheetName);
      }
    }
  }

  private static boolean containsUnquotedSheetReference(String maskedFormula, String sheetName) {
    String token = sheetName + "!";
    int fromIndex = 0;
    while (true) {
      int matchIndex = maskedFormula.indexOf(token, fromIndex);
      if (matchIndex < 0) {
        return false;
      }
      int nextIndex = matchIndex + token.length();
      if (hasReferenceBoundaryBefore(maskedFormula, matchIndex)
          && hasCellReferenceStartAfter(maskedFormula, nextIndex)) {
        return true;
      }
      fromIndex = matchIndex + 1;
    }
  }

  private static boolean hasReferenceBoundaryBefore(String formula, int index) {
    if (index == 0) {
      return true;
    }
    return switch (formula.charAt(index - 1)) {
      case '(', ',', ';', '+', '-', '*', '/', '^', '&', '=', '<', '>', ':' -> true;
      default -> false;
    };
  }

  private static boolean hasCellReferenceStartAfter(String formula, int index) {
    if (index >= formula.length()) {
      return false;
    }
    char current = formula.charAt(index);
    return current == '$' || Character.isLetter(current);
  }

  private static String maskDoubleQuotedStrings(String formula) {
    StringBuilder masked = new StringBuilder(formula.length());
    boolean insideString = false;
    int index = 0;
    while (index < formula.length()) {
      char current = formula.charAt(index);
      if (current == '"') {
        masked.append(current);
        if (insideString && index + 1 < formula.length() && formula.charAt(index + 1) == '"') {
          masked.append('"');
          index++;
        } else {
          insideString = !insideString;
        }
      } else {
        masked.append(insideString ? ' ' : current);
      }
      index++;
    }
    return masked.toString();
  }
}
