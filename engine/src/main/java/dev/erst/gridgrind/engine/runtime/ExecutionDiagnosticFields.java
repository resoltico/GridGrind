package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.Optional;

/** Resolves workbook-step and exception facts into journal-friendly diagnostic locations. */
final class ExecutionDiagnosticFields {
  private ExecutionDiagnosticFields() {}

  static ProblemContextWorkbookSurfaces.ProblemLocation locationFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep ->
          ProblemContext.mergeLocation(
              locationFor(mutationStep.action()), locationFor(mutationStep.target()));
      case AssertionStep assertionStep ->
          ProblemContext.mergeLocation(
              locationFor(assertionStep.target()), locationFor(assertionStep.assertion()));
      case InspectionStep inspectionStep -> locationFor(inspectionStep.target());
    };
  }

  static ProblemContextWorkbookSurfaces.ProblemLocation locationFor(
      WorkbookStep step, Exception exception) {
    return ProblemContext.mergeLocation(locationFor(step), locationFor(exception));
  }

  static ProblemContextWorkbookSurfaces.ProblemLocation locationFor(Exception exception) {
    return locationFor(
        new DiagnosticLocationParts(
            ExecutionExceptionDiagnosticFields.namedRangeNameFor(exception),
            ExecutionExceptionDiagnosticFields.sheetNameFor(exception),
            ExecutionExceptionDiagnosticFields.addressFor(exception),
            ExecutionExceptionDiagnosticFields.rangeFor(exception),
            ExecutionExceptionDiagnosticFields.formulaFor(exception)));
  }

  static ProblemContextWorkbookSurfaces.ProblemLocation locationFor(MutationAction action) {
    return locationFor(
        new DiagnosticLocationParts(
            ExecutionActionDiagnosticFields.namedRangeNameFor(action),
            ExecutionActionDiagnosticFields.sheetNameFor(action),
            ExecutionActionDiagnosticFields.addressFor(action),
            ExecutionActionDiagnosticFields.rangeFor(action),
            Optional.empty()));
  }

  static ProblemContextWorkbookSurfaces.ProblemLocation locationFor(Assertion assertion) {
    return ProblemContextWorkbookSurfaces.ProblemLocation.unknown();
  }

  static ProblemContextWorkbookSurfaces.ProblemLocation locationFor(Selector selector) {
    return locationFor(
        new DiagnosticLocationParts(
            ExecutionSelectorDiagnosticFields.namedRangeNameFor(selector),
            ExecutionSelectorDiagnosticFields.sheetNameFor(selector),
            ExecutionSelectorDiagnosticFields.addressFor(selector),
            ExecutionSelectorDiagnosticFields.rangeFor(selector),
            Optional.empty()));
  }

  static Optional<String> sheetNameFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep ->
          ExecutionSelectorDiagnosticFields.sheetNameFor(mutationStep.target())
              .or(() -> ExecutionActionDiagnosticFields.sheetNameFor(mutationStep.action()));
      case AssertionStep assertionStep ->
          ExecutionSelectorDiagnosticFields.sheetNameFor(assertionStep.target());
      case InspectionStep inspectionStep ->
          ExecutionSelectorDiagnosticFields.sheetNameFor(inspectionStep.target());
    };
  }

  static Optional<String> sheetNameFor(WorkbookStep step, Exception exception) {
    return sheetNameFor(step).or(() -> ExecutionExceptionDiagnosticFields.sheetNameFor(exception));
  }

  static Optional<String> addressFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep ->
          ExecutionSelectorDiagnosticFields.addressFor(mutationStep.target())
              .or(() -> ExecutionActionDiagnosticFields.addressFor(mutationStep.action()));
      case AssertionStep assertionStep ->
          ExecutionSelectorDiagnosticFields.addressFor(assertionStep.target());
      case InspectionStep inspectionStep ->
          ExecutionSelectorDiagnosticFields.addressFor(inspectionStep.target());
    };
  }

  static Optional<String> addressFor(WorkbookStep step, Exception exception) {
    return addressFor(step).or(() -> ExecutionExceptionDiagnosticFields.addressFor(exception));
  }

  static Optional<String> rangeFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep ->
          ExecutionSelectorDiagnosticFields.rangeFor(mutationStep.target())
              .or(() -> ExecutionActionDiagnosticFields.rangeFor(mutationStep.action()));
      case AssertionStep assertionStep ->
          ExecutionSelectorDiagnosticFields.rangeFor(assertionStep.target());
      case InspectionStep inspectionStep ->
          ExecutionSelectorDiagnosticFields.rangeFor(inspectionStep.target());
    };
  }

  static Optional<String> rangeFor(WorkbookStep step, Exception exception) {
    return rangeFor(step).or(() -> ExecutionExceptionDiagnosticFields.rangeFor(exception));
  }

  static Optional<String> formulaFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep ->
          ExecutionActionDiagnosticFields.formulaFor(mutationStep.action());
      case AssertionStep assertionStep ->
          ExecutionActionDiagnosticFields.formulaFor(assertionStep.assertion());
      case InspectionStep _ -> Optional.empty();
    };
  }

  static Optional<String> formulaFor(WorkbookStep step, Exception exception) {
    return formulaFor(step).or(() -> ExecutionExceptionDiagnosticFields.formulaFor(exception));
  }

  static Optional<String> namedRangeNameFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep ->
          ExecutionSelectorDiagnosticFields.namedRangeNameFor(mutationStep.target())
              .or(() -> ExecutionActionDiagnosticFields.namedRangeNameFor(mutationStep.action()));
      case AssertionStep assertionStep ->
          ExecutionSelectorDiagnosticFields.namedRangeNameFor(assertionStep.target());
      case InspectionStep inspectionStep ->
          ExecutionSelectorDiagnosticFields.namedRangeNameFor(inspectionStep.target());
    };
  }

  static Optional<String> namedRangeNameFor(WorkbookStep step, Exception exception) {
    return namedRangeNameFor(step)
        .or(() -> ExecutionExceptionDiagnosticFields.namedRangeNameFor(exception));
  }

  private static ProblemContextWorkbookSurfaces.ProblemLocation locationFor(
      DiagnosticLocationParts parts) {
    return namedRangeLocation(parts)
        .or(() -> formulaCellLocation(parts))
        .or(() -> cellLocation(parts))
        .or(() -> rangeLocation(parts))
        .or(() -> sheetLocation(parts))
        .or(() -> addressLocation(parts))
        .or(() -> bareRangeLocation(parts))
        .orElse(ProblemContextWorkbookSurfaces.ProblemLocation.unknown());
  }

  private static Optional<ProblemContextWorkbookSurfaces.ProblemLocation> namedRangeLocation(
      DiagnosticLocationParts parts) {
    if (parts.namedRange().isPresent() && parts.sheetName().isPresent()) {
      return Optional.of(
          ProblemContextWorkbookSurfaces.ProblemLocation.namedRange(
              parts.sheetName().orElseThrow(), parts.namedRange().orElseThrow()));
    }
    if (parts.namedRange().isPresent()) {
      return Optional.of(
          ProblemContextWorkbookSurfaces.ProblemLocation.namedRange(
              parts.namedRange().orElseThrow()));
    }
    return Optional.empty();
  }

  private static Optional<ProblemContextWorkbookSurfaces.ProblemLocation> formulaCellLocation(
      DiagnosticLocationParts parts) {
    if (parts.sheetName().isPresent()
        && parts.address().isPresent()
        && parts.formula().isPresent()) {
      return Optional.of(
          ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
              parts.sheetName().orElseThrow(),
              parts.address().orElseThrow(),
              parts.formula().orElseThrow()));
    }
    return Optional.empty();
  }

  private static Optional<ProblemContextWorkbookSurfaces.ProblemLocation> cellLocation(
      DiagnosticLocationParts parts) {
    if (parts.sheetName().isPresent() && parts.address().isPresent()) {
      return Optional.of(
          ProblemContextWorkbookSurfaces.ProblemLocation.cell(
              parts.sheetName().orElseThrow(), parts.address().orElseThrow()));
    }
    return Optional.empty();
  }

  private static Optional<ProblemContextWorkbookSurfaces.ProblemLocation> rangeLocation(
      DiagnosticLocationParts parts) {
    if (parts.sheetName().isPresent() && parts.range().isPresent()) {
      return Optional.of(
          ProblemContextWorkbookSurfaces.ProblemLocation.range(
              parts.sheetName().orElseThrow(), parts.range().orElseThrow()));
    }
    return Optional.empty();
  }

  private static Optional<ProblemContextWorkbookSurfaces.ProblemLocation> sheetLocation(
      DiagnosticLocationParts parts) {
    return parts.sheetName().map(ProblemContextWorkbookSurfaces.ProblemLocation::sheet);
  }

  private static Optional<ProblemContextWorkbookSurfaces.ProblemLocation> addressLocation(
      DiagnosticLocationParts parts) {
    return parts.address().map(ProblemContextWorkbookSurfaces.ProblemLocation::address);
  }

  private static Optional<ProblemContextWorkbookSurfaces.ProblemLocation> bareRangeLocation(
      DiagnosticLocationParts parts) {
    return parts.range().map(ProblemContextWorkbookSurfaces.ProblemLocation::range);
  }

  private record DiagnosticLocationParts(
      Optional<String> namedRange,
      Optional<String> sheetName,
      Optional<String> address,
      Optional<String> range,
      Optional<String> formula) {}
}
