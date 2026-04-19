package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.stream.IntStream;

/** Collects non-fatal request warnings that can be derived directly from one parsed request. */
final class GridGrindRequestWarnings {
  private GridGrindRequestWarnings() {}

  static List<RequestWarning> collect(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");

    List<MutationStep> mutationSteps =
        request.steps().stream()
            .filter(MutationStep.class::isInstance)
            .map(MutationStep.class::cast)
            .toList();
    Set<String> spacedSheetNames = sameRequestSpacedSheetNames(mutationSteps);
    if (spacedSheetNames.isEmpty()) {
      return List.of();
    }

    return IntStream.range(0, request.steps().size())
        .mapToObj(
            stepIndex -> warningFor(request.steps().get(stepIndex), stepIndex, spacedSheetNames))
        .filter(Objects::nonNull)
        .toList();
  }

  private static Set<String> sameRequestSpacedSheetNames(List<MutationStep> steps) {
    Set<String> sheetNames = new LinkedHashSet<>();
    for (MutationStep step : steps) {
      switch (step.action()) {
        case MutationAction.EnsureSheet _ ->
            addIfSpaced(
                sheetNames,
                SelectorConverter.toSheetName(
                    (dev.erst.gridgrind.contract.selector.SheetSelector.ByName) step.target()));
        case MutationAction.RenameSheet renameSheet ->
            addIfSpaced(sheetNames, renameSheet.newSheetName());
        case MutationAction.CopySheet copySheet ->
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
      MutationStep step, Set<String> spacedSheetNames, Set<String> offendingSheetNames) {
    switch (step.action()) {
      case MutationAction.SetCell setCell ->
          collectFromCellInput(setCell.value(), spacedSheetNames, offendingSheetNames);
      case MutationAction.SetRange setRange -> {
        for (List<CellInput> row : setRange.rows()) {
          for (CellInput cellInput : row) {
            collectFromCellInput(cellInput, spacedSheetNames, offendingSheetNames);
          }
        }
      }
      case MutationAction.AppendRow appendRow -> {
        for (CellInput cellInput : appendRow.values()) {
          collectFromCellInput(cellInput, spacedSheetNames, offendingSheetNames);
        }
      }
      default -> {}
    }
  }

  private static RequestWarning warningFor(
      WorkbookStep step, int stepIndex, Set<String> spacedSheetNames) {
    if (!(step instanceof MutationStep mutationStep)) {
      return null;
    }
    Set<String> offendingSheetNames = new LinkedHashSet<>();
    collectOffendingSheetNames(mutationStep, spacedSheetNames, offendingSheetNames);
    if (offendingSheetNames.isEmpty()) {
      return null;
    }
    return new RequestWarning(
        stepIndex,
        mutationStep.stepId(),
        mutationStep.action().actionType(),
        "Formula references same-request sheet names with spaces without single quotes: "
            + String.join(", ", offendingSheetNames)
            + ". Use 'Sheet Name'!A1 syntax.");
  }

  private static void collectFromCellInput(
      CellInput input, Set<String> spacedSheetNames, Set<String> offendingSheetNames) {
    if (!(input instanceof CellInput.Formula formula)) {
      return;
    }
    String formulaText =
        formula.source() instanceof TextSourceInput.Inline inline ? inline.text() : null;
    if (formulaText == null) {
      return;
    }
    String maskedFormula = maskDoubleQuotedStrings(formulaText);
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
