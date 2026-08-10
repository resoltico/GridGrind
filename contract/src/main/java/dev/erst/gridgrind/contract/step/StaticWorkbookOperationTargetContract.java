package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Explicit selector family list for an operation with static targeting. */
record StaticWorkbookOperationTargetContract(List<Class<? extends Selector>> selectorTypes)
    implements WorkbookOperationTargetContract {
  StaticWorkbookOperationTargetContract {
    Objects.requireNonNull(selectorTypes, "selectorTypes must not be null");
    selectorTypes = List.copyOf(selectorTypes);
    if (selectorTypes.isEmpty()) {
      throw new IllegalArgumentException("selectorTypes must not be empty");
    }
  }

  @Override
  public List<Class<? extends Selector>> acceptedSelectors(Object operation) {
    Objects.requireNonNull(operation, "operation must not be null");
    return selectorTypes;
  }

  @Override
  public WorkbookStepTargeting.TargetSurface targetSurface() {
    return new WorkbookStepTargeting.TargetSurface(
        SelectorJsonSupport.familyInfosFor(selectorTypes), Optional.empty());
  }
}
