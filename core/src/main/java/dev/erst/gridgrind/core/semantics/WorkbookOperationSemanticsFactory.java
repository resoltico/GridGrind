package dev.erst.gridgrind.core.semantics;

import dev.erst.gridgrind.core.metadata.Variant;
import java.util.Arrays;

/** Derives canonical effects and preconditions from colocated operation metadata. */
final class WorkbookOperationSemanticsFactory {
  private WorkbookOperationSemanticsFactory() {}

  static OperationSemantics semanticsFor(Class<? extends Record> operationType) {
    Variant metadata = operationType.getAnnotation(Variant.class);
    if (metadata == null || metadata.profile() == OperationProfile.NONE) {
      throw new IllegalArgumentException(
          "operation type has no static semantics: " + operationType);
    }
    return new OperationSemantics(
        metadata.profile().effects(),
        metadata.profile().footprint(),
        Arrays.stream(metadata.preconditions())
            .map(WorkbookOperationSemanticsFactory::precondition)
            .toList());
  }

  private static OperationPrecondition precondition(Class<? extends OperationPrecondition> type) {
    try {
      return type.getConstructor().newInstance();
    } catch (ReflectiveOperationException failure) {
      throw new IllegalStateException(
          "Operation preconditions need a public no-argument constructor: " + type, failure);
    }
  }
}
