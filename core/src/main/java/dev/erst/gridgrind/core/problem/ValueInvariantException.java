package dev.erst.gridgrind.core.problem;

import java.util.Objects;

/** Explicitly identifies a declared core value invariant, independent of a transport role. */
public non-sealed class ValueInvariantException extends IllegalArgumentException
    implements RequestProblemSource {
  private static final long serialVersionUID = 1L;
  private final RequestProblemDescriptor.Invariant problem;

  /** Creates a violation only at the owner of the declared value constraint. */
  public ValueInvariantException(RequestProblemDescriptor.Invariant problem) {
    super(
        GridGrindRequestProblemSupport.message(
            Objects.requireNonNull(problem, "problem must not be null")));
    this.problem = problem;
  }

  @Override
  public final RequestProblemDescriptor.Invariant requestProblem() {
    return problem;
  }
}
