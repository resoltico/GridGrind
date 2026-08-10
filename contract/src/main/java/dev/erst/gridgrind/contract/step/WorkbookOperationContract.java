package dev.erst.gridgrind.contract.step;

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
  private final WorkbookOperationTargetContract targetSelectorContract;

  WorkbookOperationContract(WorkbookOperationTargetContract targetSelectorContract) {
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

  @SuppressWarnings("unchecked")
  private static Class<? extends Selector>[] toArray(
      java.util.List<Class<? extends Selector>> values) {
    return values.toArray(new Class[0]);
  }
}
