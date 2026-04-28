package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.Optional;

/** Resolves workbook-step and exception facts into journal-friendly diagnostic locations. */
final class ExecutionDiagnosticFields {
  private ExecutionDiagnosticFields() {}

  static ProblemContext.ProblemLocation locationFor(WorkbookStep step) {
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

  static ProblemContext.ProblemLocation locationFor(WorkbookStep step, Exception exception) {
    return ProblemContext.mergeLocation(locationFor(step), locationFor(exception));
  }

  static ProblemContext.ProblemLocation locationFor(Exception exception) {
    Optional<String> namedRange = ExecutionExceptionDiagnosticFields.namedRangeNameFor(exception);
    Optional<String> sheetName = ExecutionExceptionDiagnosticFields.sheetNameFor(exception);
    Optional<String> address = ExecutionExceptionDiagnosticFields.addressFor(exception);
    Optional<String> range = ExecutionExceptionDiagnosticFields.rangeFor(exception);
    Optional<String> formula = ExecutionExceptionDiagnosticFields.formulaFor(exception);
    if (namedRange.isPresent() && sheetName.isPresent()) {
      return ProblemContext.ProblemLocation.namedRange(
          sheetName.orElseThrow(), namedRange.orElseThrow());
    }
    if (namedRange.isPresent()) {
      return ProblemContext.ProblemLocation.namedRange(namedRange.orElseThrow());
    }
    if (sheetName.isPresent() && address.isPresent() && formula.isPresent()) {
      return ProblemContext.ProblemLocation.formulaCell(
          sheetName.orElseThrow(), address.orElseThrow(), formula.orElseThrow());
    }
    if (sheetName.isPresent() && address.isPresent()) {
      return ProblemContext.ProblemLocation.cell(sheetName.orElseThrow(), address.orElseThrow());
    }
    if (sheetName.isPresent()) {
      return ProblemContext.ProblemLocation.sheet(sheetName.orElseThrow());
    }
    if (address.isPresent()) {
      return ProblemContext.ProblemLocation.address(address.orElseThrow());
    }
    if (range.isPresent()) {
      return ProblemContext.ProblemLocation.range(range.orElseThrow());
    }
    return ProblemContext.ProblemLocation.unknown();
  }

  static ProblemContext.ProblemLocation locationFor(MutationAction action) {
    Optional<String> namedRange = ExecutionActionDiagnosticFields.namedRangeNameFor(action);
    Optional<String> sheetName = ExecutionActionDiagnosticFields.sheetNameFor(action);
    Optional<String> address = ExecutionActionDiagnosticFields.addressFor(action);
    Optional<String> range = ExecutionActionDiagnosticFields.rangeFor(action);
    if (namedRange.isPresent() && sheetName.isPresent()) {
      return ProblemContext.ProblemLocation.namedRange(
          sheetName.orElseThrow(), namedRange.orElseThrow());
    }
    if (namedRange.isPresent()) {
      return ProblemContext.ProblemLocation.namedRange(namedRange.orElseThrow());
    }
    if (sheetName.isPresent() && address.isPresent()) {
      return ProblemContext.ProblemLocation.cell(sheetName.orElseThrow(), address.orElseThrow());
    }
    if (range.isPresent()) {
      return sheetName.isPresent()
          ? ProblemContext.ProblemLocation.range(sheetName.orElseThrow(), range.orElseThrow())
          : ProblemContext.ProblemLocation.range(range.orElseThrow());
    }
    return ProblemContext.ProblemLocation.unknown();
  }

  static ProblemContext.ProblemLocation locationFor(Assertion assertion) {
    return ProblemContext.ProblemLocation.unknown();
  }

  static ProblemContext.ProblemLocation locationFor(Selector selector) {
    Optional<String> namedRange = ExecutionSelectorDiagnosticFields.namedRangeNameFor(selector);
    Optional<String> sheetName = ExecutionSelectorDiagnosticFields.sheetNameFor(selector);
    Optional<String> address = ExecutionSelectorDiagnosticFields.addressFor(selector);
    Optional<String> range = ExecutionSelectorDiagnosticFields.rangeFor(selector);
    if (namedRange.isPresent() && sheetName.isPresent()) {
      return ProblemContext.ProblemLocation.namedRange(
          sheetName.orElseThrow(), namedRange.orElseThrow());
    }
    if (namedRange.isPresent()) {
      return ProblemContext.ProblemLocation.namedRange(namedRange.orElseThrow());
    }
    if (sheetName.isPresent() && address.isPresent()) {
      return ProblemContext.ProblemLocation.cell(sheetName.orElseThrow(), address.orElseThrow());
    }
    if (sheetName.isPresent() && range.isPresent()) {
      return ProblemContext.ProblemLocation.range(sheetName.orElseThrow(), range.orElseThrow());
    }
    if (sheetName.isPresent()) {
      return ProblemContext.ProblemLocation.sheet(sheetName.orElseThrow());
    }
    return ProblemContext.ProblemLocation.unknown();
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

  static Optional<String> sheetNameFor(Exception exception) {
    return ExecutionExceptionDiagnosticFields.sheetNameFor(exception);
  }

  static Optional<String> addressFor(Exception exception) {
    return ExecutionExceptionDiagnosticFields.addressFor(exception);
  }

  static Optional<String> rangeFor(Exception exception) {
    return ExecutionExceptionDiagnosticFields.rangeFor(exception);
  }

  static Optional<String> formulaFor(Exception exception) {
    return ExecutionExceptionDiagnosticFields.formulaFor(exception);
  }

  static Optional<String> namedRangeNameFor(Exception exception) {
    return ExecutionExceptionDiagnosticFields.namedRangeNameFor(exception);
  }

  static Optional<String> sheetNameFor(MutationAction action) {
    return ExecutionActionDiagnosticFields.sheetNameFor(action);
  }

  static Optional<String> addressFor(MutationAction action) {
    return ExecutionActionDiagnosticFields.addressFor(action);
  }

  static Optional<String> rangeFor(MutationAction action) {
    return ExecutionActionDiagnosticFields.rangeFor(action);
  }

  static Optional<String> formulaFor(MutationAction action) {
    return ExecutionActionDiagnosticFields.formulaFor(action);
  }

  static Optional<String> formulaFor(Assertion assertion) {
    return ExecutionActionDiagnosticFields.formulaFor(assertion);
  }

  static Optional<String> namedRangeNameFor(MutationAction action) {
    return ExecutionActionDiagnosticFields.namedRangeNameFor(action);
  }

  static Optional<String> sheetNameFor(Selector selector) {
    return ExecutionSelectorDiagnosticFields.sheetNameFor(selector);
  }

  static Optional<String> addressFor(Selector selector) {
    return ExecutionSelectorDiagnosticFields.addressFor(selector);
  }

  static Optional<String> rangeFor(Selector selector) {
    return ExecutionSelectorDiagnosticFields.rangeFor(selector);
  }

  static Optional<String> namedRangeNameFor(Selector selector) {
    return ExecutionSelectorDiagnosticFields.namedRangeNameFor(selector);
  }

  static Optional<String> singleSheetName(CellSelector selector) {
    return ExecutionSelectorDiagnosticFields.singleSheetName(selector);
  }

  static Optional<String> singleSheetName(RangeSelector selector) {
    return ExecutionSelectorDiagnosticFields.singleSheetName(selector);
  }

  static Optional<String> singleSheetName(NamedRangeSelector selector) {
    return ExecutionSelectorDiagnosticFields.singleSheetName(selector);
  }

  static Optional<String> singleSheetName(NamedRangeSelector.Ref selector) {
    return ExecutionSelectorDiagnosticFields.singleSheetName(selector);
  }

  static Optional<String> singleNamedRangeName(NamedRangeSelector selector) {
    return ExecutionSelectorDiagnosticFields.singleNamedRangeName(selector);
  }

  static Optional<String> singleNamedRangeName(NamedRangeSelector.Ref selector) {
    return ExecutionSelectorDiagnosticFields.singleNamedRangeName(selector);
  }
}
