package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Typed request problems shared across request decode, doctor, and reporting surfaces. */
public sealed interface RequestProblemDescriptor
    permits RequestProblemDescriptor.Shape, RequestProblemDescriptor.Invariant {
  /** Structured JSON path when the problem pinpoints one authored field. */
  Optional<String> jsonPath();

  /** Structural request-shape problem. */
  sealed interface Shape extends RequestProblemDescriptor
      permits MissingRequiredField,
          MissingTypeDiscriminator,
          UnknownField,
          UnknownTypeValue,
          UnsupportedValue,
          ExplicitNullField,
          ActionableShapeMessage,
          MessageShape {}

  /** Semantic request problem after JSON successfully bound. */
  sealed interface Invariant extends RequestProblemDescriptor
      permits DuplicateStepId,
          NonXlsxPath,
          FieldValidationProblem,
          ActionableInvariantMessage,
          MessageInvariant {}
}
