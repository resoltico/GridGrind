package dev.erst.gridgrind.contract.assertion;

import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.WorkbookOperationContracts;
import java.util.Objects;

/** First-class verification contract evaluated against a workbook target. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
public sealed interface Assertion
    permits PresenceAssertion,
        CellAssertion,
        WorkbookFactAssertion,
        AnalysisAssertion,
        CompositeAssertion {

  /** Stable SCREAMING_SNAKE_CASE discriminator mirrored in catalog and result surfaces. */
  default String assertionType() {
    return GridGrindProtocolTypeNames.assertionTypeName(getClass().asSubclass(Assertion.class));
  }

  /** Returns the selector types accepted by one assertion instance. */
  static Class<? extends Selector>[] allowedTargetTypes(Assertion assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    return WorkbookOperationContracts.targetSelectorsFor(assertion);
  }
}
