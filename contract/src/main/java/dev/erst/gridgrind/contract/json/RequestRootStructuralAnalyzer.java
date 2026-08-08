package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.catalog.GridGrindProtocolContractSupport;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/** Analyses the request envelope and coordinates independently bound root fragments. */
final class RequestRootStructuralAnalyzer {
  private RequestRootStructuralAnalyzer() {}

  static RequestAnalysis analyze(RequestSyntaxParseResult parsed) {
    List<RequestStructuralProblem> problems = new ArrayList<>(parsed.problems());
    List<RequestBindingFailure> bindingFailures = new ArrayList<>();
    RequestDiagnosticRedactor diagnosticRedactor =
        RequestDiagnosticRedactor.forRequestType(WorkbookPlan.class);
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
      return new RequestAnalysis(
          emptyFragments(),
          problems,
          bindingFailures,
          diagnosticRedactor,
          Optional.of(parsed.root()));
    }

    RequestObjectMembers.Index members = RequestObjectMembers.index(root);
    GridGrindProtocolContractSupport.EffectiveObjectContract rootContract =
        GridGrindProtocolContractSupport.effectiveObjectContract(WorkbookPlan.class);
    RequestObjectMembers.collectFields(
        members, "", rootContract.requiredFields(), rootContract.optionalFields(), problems);
    validateRootScalars(members, problems);
    validateRootMembers(members, "source", WorkbookPlan.WorkbookSource.class, problems);
    validateRootMembers(members, "persistence", WorkbookPlan.WorkbookPersistence.class, problems);
    validateRootMembers(members, "execution", ExecutionPolicyInput.class, problems);
    validateRootMembers(members, "formulaEnvironment", FormulaEnvironmentInput.class, problems);
    List<RequestJsonMember> stepMembers = members.membersNamed("steps");
    RequestStepFragmentCollector.validateAll(stepMembers, problems);

    return new RequestAnalysis(
        new RequestBoundFragments(
            new RequestBoundRoot(
                RequestFragmentBinder.bindMember(
                    members.member("protocolVersion"),
                    GridGrindProtocolVersion.class,
                    "protocolVersion",
                    problems,
                    bindingFailures,
                    diagnosticRedactor),
                RequestFragmentBinder.bindMember(
                    members.member("planId"),
                    String.class,
                    "planId",
                    problems,
                    bindingFailures,
                    diagnosticRedactor),
                RequestFragmentBinder.bindMember(
                    members.member("source"),
                    WorkbookPlan.WorkbookSource.class,
                    "source",
                    problems,
                    bindingFailures,
                    diagnosticRedactor),
                RequestFragmentBinder.bindMember(
                    members.member("persistence"),
                    WorkbookPlan.WorkbookPersistence.class,
                    "persistence",
                    problems,
                    bindingFailures,
                    diagnosticRedactor),
                bindOrDefault(
                    members.member("execution"),
                    ExecutionPolicyInput.class,
                    "execution",
                    ExecutionPolicyInput.defaults(),
                    problems,
                    bindingFailures,
                    diagnosticRedactor),
                bindOrDefault(
                    members.member("formulaEnvironment"),
                    FormulaEnvironmentInput.class,
                    "formulaEnvironment",
                    FormulaEnvironmentInput.empty(),
                    problems,
                    bindingFailures,
                    diagnosticRedactor)),
            RequestStepFragmentCollector.bind(
                members.member("steps"), problems, bindingFailures, diagnosticRedactor)),
        problems,
        bindingFailures,
        diagnosticRedactor,
        Optional.of(parsed.root()));
  }

  private static void validateRootScalars(
      RequestObjectMembers.Index root, List<RequestStructuralProblem> problems) {
    for (RequestJsonMember member : root.membersNamed("protocolVersion")) {
      RequestNodeValidator.validateEnum(
          member.value(),
          GridGrindProtocolVersion.class,
          "protocolVersion",
          member.nameByteOffset(),
          problems);
    }
    validateRootMembers(root, "planId", String.class, problems);
  }

  private static void validateRootMembers(
      RequestObjectMembers.Index root,
      String fieldName,
      Type expectedType,
      List<RequestStructuralProblem> problems) {
    for (RequestJsonMember member : root.membersNamed(fieldName)) {
      RequestNodeValidator.validateNode(
          member.value(), expectedType, fieldName, member.nameByteOffset(), problems);
    }
  }

  private static <T> Optional<T> bindOrDefault(
      Optional<RequestJsonMember> member,
      Class<T> type,
      String jsonPath,
      T defaultValue,
      List<RequestStructuralProblem> structuralProblems,
      List<RequestBindingFailure> bindingFailures,
      RequestDiagnosticRedactor diagnosticRedactor) {
    return member.isPresent()
        ? RequestFragmentBinder.bindMember(
            member, type, jsonPath, structuralProblems, bindingFailures, diagnosticRedactor)
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
