package dev.erst.gridgrind.core.plan;

import java.util.Objects;

/** A mutation with materialized non-secret operands and its original diagnostic identity. */
public record PreparedMutationStep(MutationStep authoredStep, PreparedMutationAction action)
    implements PreparedStep {
  public PreparedMutationStep {
    Objects.requireNonNull(authoredStep, "authoredStep must not be null");
    Objects.requireNonNull(action, "action must not be null");
    boolean matches =
        switch (action) {
          case ResolvedMutationAction resolved ->
              resolved.authoredType().isInstance(authoredStep.action());
          case WorkbookMutationAction.SetSheetProtection protection ->
              protection.equals(authoredStep.action());
          case WorkbookMutationAction.SetWorkbookProtection protection ->
              protection.equals(authoredStep.action());
        };
    if (!matches)
      throw new IllegalArgumentException("Prepared action must match its authored mutation");
  }
}
