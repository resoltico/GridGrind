package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/** Validates and independently binds authored request steps. */
final class RequestStepFragmentCollector {
  private RequestStepFragmentCollector() {}

  /** Validates every authored {@code steps} member before the first occurrence is bound. */
  static void validateAll(
      List<RequestJsonMember> stepsMembers, List<RequestStructuralProblem> problems) {
    for (RequestJsonMember stepsMember : stepsMembers) {
      validateStepsMember(stepsMember, problems);
    }
  }

  /** Binds the first {@code steps} member only when that outer field has no structural problem. */
  static Optional<List<RequestBoundFragments.Step>> bind(
      Optional<RequestJsonMember> stepsMember,
      List<RequestStructuralProblem> structuralProblems,
      List<RequestBindingFailure> bindingFailures,
      RequestDiagnosticRedactor diagnosticRedactor) {
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
                  element,
                  WorkbookStep.class,
                  stepPath,
                  structuralProblems,
                  bindingFailures,
                  diagnosticRedactor)));
    }
    return Optional.of(List.copyOf(boundSteps));
  }

  private static void validateStepsMember(
      RequestJsonMember stepsField, List<RequestStructuralProblem> problems) {
    if (stepsField.value() instanceof RequestJsonNull) {
      problems.add(new RequestExplicitNullField("steps", stepsField.nameByteOffset()));
      return;
    }
    if (!(stepsField.value() instanceof RequestJsonArray steps)) {
      problems.add(
          new RequestMalformedScalar("steps", "a JSON array", stepsField.nameByteOffset()));
      return;
    }
    for (int index = 0; index < steps.elements().size(); index++) {
      validate(steps.elements().get(index), "steps[" + index + "]", problems);
    }
  }

  private static void validate(
      RequestJsonNode node, String jsonPath, List<RequestStructuralProblem> problems) {
    if (node instanceof RequestJsonNull) {
      problems.add(new RequestExplicitNullField(jsonPath, node.byteOffset()));
      return;
    }
    if (!(node instanceof RequestJsonObject step)) {
      problems.add(new RequestMalformedScalar(jsonPath, "a JSON object", node.byteOffset()));
      return;
    }
    RequestObjectMembers.Index members = RequestObjectMembers.index(step);
    RequestObjectMembers.collectFields(
        members,
        jsonPath,
        List.of("stepId", "target"),
        List.of("action", "assertion", "query"),
        problems);
    validateMembers(members, "stepId", String.class, jsonPath, problems);
    validatePayload(members, jsonPath, step.byteOffset(), problems);
    validateMembers(members, "target", Selector.class, jsonPath, problems);
  }

  private static void validatePayload(
      RequestObjectMembers.Index members,
      String jsonPath,
      long diagnosticByteOffset,
      List<RequestStructuralProblem> problems) {
    List<String> payloadFields =
        List.of("action", "assertion", "query").stream()
            .filter(field -> members.member(field).isPresent())
            .toList();
    for (String payloadField : payloadFields) {
      validateMembers(members, payloadField, payloadType(payloadField), jsonPath, problems);
    }
    if (payloadFields.size() != 1) {
      problems.add(
          new RequestMalformedScalar(
              jsonPath,
              "an object containing exactly one of action, assertion, or query",
              diagnosticByteOffset));
      return;
    }
  }

  private static void validateMembers(
      RequestObjectMembers.Index object,
      String fieldName,
      Class<?> expectedType,
      String parentPath,
      List<RequestStructuralProblem> problems) {
    String jsonPath = RequestObjectMembers.childPath(parentPath, fieldName);
    for (RequestJsonMember member : object.membersNamed(fieldName)) {
      RequestNodeValidator.validateNode(
          member.value(), expectedType, jsonPath, member.nameByteOffset(), problems);
    }
  }

  static Class<?> payloadType(String payloadField) {
    return switch (payloadField) {
      case "action" -> MutationAction.class;
      case "assertion" -> Assertion.class;
      case "query" -> dev.erst.gridgrind.contract.query.InspectionQuery.class;
      default -> throw new IllegalStateException("Unknown step payload field: " + payloadField);
    };
  }
}
