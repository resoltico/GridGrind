package dev.erst.gridgrind.json.intake;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.RequestBindingFailure;
import dev.erst.gridgrind.core.problem.RequestBoundFragments;
import dev.erst.gridgrind.core.problem.RequestBoundRoot;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.problem.RequestInvalidEncoding;
import dev.erst.gridgrind.core.problem.RequestInvalidJson;
import dev.erst.gridgrind.core.problem.RequestMalformedScalar;
import dev.erst.gridgrind.core.problem.RequestStructuralProblem;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/** Analyses the request envelope and coordinates independently bound root fragments. */
final class RequestRootStructuralAnalyzer {
  private RequestRootStructuralAnalyzer() {}

  static RequestIntake analyze(RequestSyntaxParseResult parsed) {
    List<RequestStructuralProblem> problems = new ArrayList<>(parsed.problems());
    List<RequestBindingFailure> bindingFailures = new ArrayList<>();
    if (!(parsed.root() instanceof RequestJsonObject root)) {
      if (parsed.problems().stream()
          .noneMatch(
              problem ->
                  problem instanceof RequestInvalidEncoding
                      || problem instanceof RequestInvalidJson)) {
        problems.add(
            new RequestMalformedScalar(
                "request", "a JSON object at the root", parsed.root().byteOffset()));
      }
      return RequestIntakeBinder.bind(
          emptyFragments(),
          problems,
          bindingFailures,
          Optional.of(parsed.root()),
          parsed.warnings());
    }

    RequestObjectMembers.Index members = RequestObjectMembers.index(root);
    RequestSchemaStructure.validate(root, problems);

    return RequestIntakeBinder.bind(
        new RequestBoundFragments(
            new RequestBoundRoot(
                RequestFragmentBinder.bindMember(
                    members.member("protocolVersion"),
                    ProtocolVersion.class,
                    "protocolVersion",
                    problems,
                    bindingFailures),
                RequestFragmentBinder.bindMember(
                    members.member("planId"),
                    dev.erst.gridgrind.core.vocabulary.PlanId.class,
                    "planId",
                    problems,
                    bindingFailures),
                RequestFragmentBinder.bindMember(
                    members.member("source"),
                    WorkbookPlan.WorkbookSource.class,
                    "source",
                    problems,
                    bindingFailures),
                RequestFragmentBinder.bindMember(
                    members.member("persistence"),
                    WorkbookPlan.WorkbookPersistence.class,
                    "persistence",
                    problems,
                    bindingFailures),
                bindOrDefault(
                    members.member("execution"),
                    ExecutionPolicySpec.class,
                    "execution",
                    ExecutionPolicySpec.defaults(),
                    problems,
                    bindingFailures),
                bindOrDefault(
                    members.member("formulaEnvironment"),
                    FormulaEnvironmentSpec.class,
                    "formulaEnvironment",
                    FormulaEnvironmentSpec.empty(),
                    problems,
                    bindingFailures)),
            RequestStepFragmentCollector.bind(members.member("steps"), problems, bindingFailures)),
        problems,
        bindingFailures,
        Optional.of(parsed.root()),
        parsed.warnings());
  }

  private static <T> Optional<T> bindOrDefault(
      Optional<RequestJsonMember> member,
      Class<T> type,
      String jsonPath,
      T defaultValue,
      List<RequestStructuralProblem> structuralProblems,
      List<RequestBindingFailure> bindingFailures) {
    return member.isPresent()
        ? RequestFragmentBinder.bindMember(
            member, type, jsonPath, structuralProblems, bindingFailures)
        : Optional.of(defaultValue);
  }

  private static RequestBoundFragments emptyFragments() {
    return new RequestBoundFragments(
        new RequestBoundRoot(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty()),
        Optional.empty());
  }
}
