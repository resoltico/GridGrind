package dev.erst.gridgrind.json.intake;

import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.RequestBindingFailure;
import dev.erst.gridgrind.core.problem.RequestBoundFragments;
import dev.erst.gridgrind.core.problem.RequestStructuralProblem;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/** Independently binds shape-validated authored steps without discarding valid siblings. */
final class RequestStepFragmentCollector {
  private RequestStepFragmentCollector() {}

  /** Binds the first {@code steps} member only when that outer field has no structural problem. */
  static Optional<List<RequestBoundFragments.Step>> bind(
      Optional<RequestJsonMember> stepsMember,
      List<RequestStructuralProblem> structuralProblems,
      List<RequestBindingFailure> bindingFailures) {
    if (stepsMember.isEmpty() || RequestFragmentBinder.hasProblemAt("steps", structuralProblems)) {
      return Optional.empty();
    }
    RequestJsonMember stepsField = stepsMember.orElseThrow();
    if (!(stepsField.value() instanceof RequestJsonArray steps)) {
      return Optional.empty();
    }
    List<RequestBoundFragments.Step> boundSteps = new ArrayList<>(steps.elements().size());
    for (int index = 0; index < steps.elements().size(); index++) {
      RequestJsonNode element = steps.elements().get(index);
      String stepPath = "steps[" + index + "]";
      boundSteps.add(
          new RequestBoundFragments.Step(
              index,
              RequestFragmentBinder.bindNode(
                  element, WorkbookStep.class, stepPath, structuralProblems, bindingFailures)));
    }
    return Optional.of(List.copyOf(boundSteps));
  }
}
