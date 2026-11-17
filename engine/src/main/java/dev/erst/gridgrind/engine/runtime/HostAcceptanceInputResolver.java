package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import java.util.ArrayList;
import java.util.List;

/** Materializes host acceptance operands under the same read authority and invocation budget. */
final class HostAcceptanceInputResolver {
  private HostAcceptanceInputResolver() {}

  static void resolve(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    if (!(bindings.acceptancePolicy() instanceof AcceptancePolicy.RequireAll policy)) return;
    List<AcceptancePolicy.Requirement> requirements = new ArrayList<>();
    int before = problems.size();
    for (int index = 0; index < policy.requirements().size(); index++) {
      var requirement = policy.requirements().get(index);
      var located =
          bindings.withInputResolutionOrigins(
              InputResolutionOrigins.forValue(
                  requirement, "grant.acceptancePolicy.requirements[" + index + "]"));
      try {
        requirements.add(
            switch (requirement) {
              case AcceptancePolicy.Requirement.TerminalAssertions terminal -> {
                List<AssertionStep> assertions = new ArrayList<>();
                for (var assertion : terminal.assertions()) {
                  assertions.add(
                      (AssertionStep) SourceBackedPlanResolver.resolveStep(assertion, located));
                }
                yield new AcceptancePolicy.Requirement.TerminalAssertions(assertions);
              }
              case AcceptancePolicy.Requirement.PreserveInspectionFacts facts -> {
                List<InspectionStep> inspections = new ArrayList<>();
                for (var inspection : facts.inspections()) {
                  inspections.add(
                      (InspectionStep) SourceBackedPlanResolver.resolveStep(inspection, located));
                }
                yield new AcceptancePolicy.Requirement.PreserveInspectionFacts(inspections);
              }
              case AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts parts -> parts;
              case AcceptancePolicy.Requirement.RequireCalculation calculation -> calculation;
            });
      } catch (InputResolutionBatchException failure) {
        for (var detail : failure.failures()) {
          problems.add(
              GridGrindProblems.fromException(
                  detail.exception(), RequestPreflight.resolveInputsContext(request, detail)));
        }
      } catch (java.io.IOException | RuntimeException failure) {
        problems.add(
            GridGrindProblems.fromException(
                failure,
                RequestPreflight.resolveInputsContext(
                    request, InputResolutionFailure.unlocated(failure))));
      }
    }
    if (problems.size() == before) {
      bindings.resolvedAcceptancePolicy(new AcceptancePolicy.RequireAll(requirements));
    }
  }
}
