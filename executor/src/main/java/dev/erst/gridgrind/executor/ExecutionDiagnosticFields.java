package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;

/** Resolves workbook-step and exception facts into journal-friendly diagnostic fields. */
final class ExecutionDiagnosticFields {
  private ExecutionDiagnosticFields() {}

  static String sheetNameFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep -> {
        String fromAction = sheetNameFor(mutationStep.action());
        yield fromAction != null ? fromAction : sheetNameFor(mutationStep.target());
      }
      case AssertionStep assertionStep -> sheetNameFor(assertionStep.target());
      case InspectionStep inspectionStep -> sheetNameFor(inspectionStep.target());
    };
  }

  static String sheetNameFor(WorkbookStep step, Exception exception) {
    String fromStep = sheetNameFor(step);
    return fromStep != null ? fromStep : sheetNameFor(exception);
  }

  static String addressFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep -> {
        String fromAction = addressFor(mutationStep.action());
        yield fromAction != null ? fromAction : addressFor(mutationStep.target());
      }
      case AssertionStep assertionStep -> addressFor(assertionStep.target());
      case InspectionStep inspectionStep -> addressFor(inspectionStep.target());
    };
  }

  static String addressFor(WorkbookStep step, Exception exception) {
    String fromStep = addressFor(step);
    return fromStep != null ? fromStep : addressFor(exception);
  }

  static String rangeFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep -> {
        String fromAction = rangeFor(mutationStep.action());
        yield fromAction != null ? fromAction : rangeFor(mutationStep.target());
      }
      case AssertionStep assertionStep -> rangeFor(assertionStep.target());
      case InspectionStep inspectionStep -> rangeFor(inspectionStep.target());
    };
  }

  static String rangeFor(WorkbookStep step, Exception exception) {
    String fromStep = rangeFor(step);
    return fromStep != null ? fromStep : rangeFor(exception);
  }

  static String formulaFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep -> formulaFor(mutationStep.action());
      case AssertionStep assertionStep -> formulaFor(assertionStep.assertion());
      case InspectionStep _ -> null;
    };
  }

  static String formulaFor(WorkbookStep step, Exception exception) {
    String fromStep = formulaFor(step);
    return fromStep != null ? fromStep : formulaFor(exception);
  }

  static String namedRangeNameFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep -> {
        String fromAction = namedRangeNameFor(mutationStep.action());
        yield fromAction != null ? fromAction : namedRangeNameFor(mutationStep.target());
      }
      case AssertionStep assertionStep -> namedRangeNameFor(assertionStep.target());
      case InspectionStep inspectionStep -> namedRangeNameFor(inspectionStep.target());
    };
  }

  static String namedRangeNameFor(WorkbookStep step, Exception exception) {
    String fromStep = namedRangeNameFor(step);
    return fromStep != null ? fromStep : namedRangeNameFor(exception);
  }

  static String sheetNameFor(Exception exception) {
    return ExecutionExceptionDiagnosticFields.sheetNameFor(exception).orElse(null);
  }

  static String addressFor(Exception exception) {
    return ExecutionExceptionDiagnosticFields.addressFor(exception).orElse(null);
  }

  static String rangeFor(Exception exception) {
    return ExecutionExceptionDiagnosticFields.rangeFor(exception).orElse(null);
  }

  static String formulaFor(Exception exception) {
    return ExecutionExceptionDiagnosticFields.formulaFor(exception).orElse(null);
  }

  static String namedRangeNameFor(Exception exception) {
    return ExecutionExceptionDiagnosticFields.namedRangeNameFor(exception).orElse(null);
  }

  static String sheetNameFor(MutationAction action) {
    return ExecutionActionDiagnosticFields.sheetNameFor(action).orElse(null);
  }

  static String addressFor(MutationAction action) {
    return ExecutionActionDiagnosticFields.addressFor(action).orElse(null);
  }

  static String rangeFor(MutationAction action) {
    return ExecutionActionDiagnosticFields.rangeFor(action).orElse(null);
  }

  static String formulaFor(MutationAction action) {
    return ExecutionActionDiagnosticFields.formulaFor(action).orElse(null);
  }

  static String formulaFor(Assertion assertion) {
    return ExecutionActionDiagnosticFields.formulaFor(assertion).orElse(null);
  }

  static String namedRangeNameFor(MutationAction action) {
    return ExecutionActionDiagnosticFields.namedRangeNameFor(action).orElse(null);
  }

  static String sheetNameFor(Selector selector) {
    return ExecutionSelectorDiagnosticFields.sheetNameFor(selector).orElse(null);
  }

  static String addressFor(Selector selector) {
    return ExecutionSelectorDiagnosticFields.addressFor(selector).orElse(null);
  }

  static String rangeFor(Selector selector) {
    return ExecutionSelectorDiagnosticFields.rangeFor(selector).orElse(null);
  }

  static String namedRangeNameFor(Selector selector) {
    return ExecutionSelectorDiagnosticFields.namedRangeNameFor(selector).orElse(null);
  }

  static String singleSheetName(CellSelector selector) {
    return ExecutionSelectorDiagnosticFields.singleSheetName(selector).orElse(null);
  }

  static String singleSheetName(RangeSelector selector) {
    return ExecutionSelectorDiagnosticFields.singleSheetName(selector).orElse(null);
  }

  static String singleSheetName(NamedRangeSelector selector) {
    return ExecutionSelectorDiagnosticFields.singleSheetName(selector).orElse(null);
  }

  static String singleSheetName(NamedRangeSelector.Ref selector) {
    return ExecutionSelectorDiagnosticFields.singleSheetName(selector).orElse(null);
  }

  static String singleNamedRangeName(NamedRangeSelector selector) {
    return ExecutionSelectorDiagnosticFields.singleNamedRangeName(selector).orElse(null);
  }

  static String singleNamedRangeName(NamedRangeSelector.Ref selector) {
    return ExecutionSelectorDiagnosticFields.singleNamedRangeName(selector).orElse(null);
  }
}
