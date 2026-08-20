package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.catalog.GridGrindExecutionModeMetadata;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadataSupport;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import java.util.Objects;
import java.util.Optional;

/**
 * The complete static targeting contract for one concrete workbook operation.
 *
 * <p>Step records deliberately model one syntactically valid operation-target pair even when the
 * pair is not executable. This contract owns the later static compatibility decision so request
 * analysis can retain independently bound sibling fragments.
 */
final class WorkbookOperationContract {
  private static final GridGrindExecutionModeMetadata.EventReadMode EVENT_READ =
      GridGrindExecutionModeMetadata.eventRead();
  private static final GridGrindExecutionModeMetadata.StreamingWriteMode STREAMING_WRITE =
      GridGrindExecutionModeMetadata.streamingWrite();

  private final Class<? extends Record> operationType;
  private final WorkbookOperationTargetContract targetSelectorContract;

  WorkbookOperationContract(
      Class<? extends Record> operationType,
      WorkbookOperationTargetContract targetSelectorContract) {
    this.operationType = Objects.requireNonNull(operationType, "operationType must not be null");
    Objects.requireNonNull(targetSelectorContract, "targetSelectorContract must not be null");
    this.targetSelectorContract = targetSelectorContract;
  }

  /** Returns the target families published for this operation's catalog entry. */
  WorkbookStepTargeting.TargetSurface targetSurface() {
    return targetSelectorContract.targetSurface();
  }

  /** Returns the incompatibility message when this operation cannot act on {@code target}. */
  Optional<String> targetViolation(Object operation, Selector target, String operationId) {
    Objects.requireNonNull(operation, "operation must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(operationId, "operationId must not be null");
    var accepted = targetSelectorContract.acceptedSelectors(operation);
    for (Class<? extends Selector> selectorType : accepted) {
      if (selectorType.isInstance(target)) {
        return Optional.empty();
      }
    }
    return Optional.of(
        operationId
            + " requires target type "
            + SelectorTargetingSupport.humanTargetTypes(toArray(accepted))
            + " but got "
            + SelectorJsonSupport.typeIdFor(target.getClass()));
  }

  Class<? extends Selector>[] acceptedSelectors(Object operation) {
    return toArray(targetSelectorContract.acceptedSelectors(operation));
  }

  /** Returns the execution-mode incompatibility for this operation when one exists. */
  Optional<String> executionModeViolation(ExecutionModeInput executionMode) {
    Objects.requireNonNull(executionMode, "executionMode must not be null");
    return switch (executionMode) {
      case ExecutionModeInput.FullXssf _ -> Optional.empty();
      case ExecutionModeInput.EventRead _ -> eventReadViolation();
      case ExecutionModeInput.StreamingWrite _ -> streamingWriteViolation();
    };
  }

  private Optional<String> eventReadViolation() {
    if (!InspectionQuery.class.isAssignableFrom(operationType)) {
      return Optional.of(EVENT_READ.unsupportedStepMessage(stepKind()));
    }
    @SuppressWarnings("unchecked")
    Class<? extends InspectionQuery> queryType = (Class<? extends InspectionQuery>) operationType;
    return EVENT_READ.allowedQueries().contains(queryType)
        ? Optional.empty()
        : Optional.of(
            EVENT_READ.unsupportedQueryMessage(
                ProtocolTypeMetadataSupport.requiredTypeId(operationType)));
  }

  private Optional<String> streamingWriteViolation() {
    if (!MutationAction.class.isAssignableFrom(operationType)) {
      return Optional.empty();
    }
    @SuppressWarnings("unchecked")
    Class<? extends MutationAction> actionType = (Class<? extends MutationAction>) operationType;
    return STREAMING_WRITE.allowedActions().contains(actionType)
        ? Optional.empty()
        : Optional.of(
            STREAMING_WRITE.unsupportedActionMessage(
                ProtocolTypeMetadataSupport.requiredTypeId(operationType)));
  }

  private String stepKind() {
    return Assertion.class.isAssignableFrom(operationType) ? "ASSERTION" : "MUTATION";
  }

  @SuppressWarnings("unchecked")
  private static Class<? extends Selector>[] toArray(
      java.util.List<Class<? extends Selector>> values) {
    return values.toArray(new Class[0]);
  }
}
