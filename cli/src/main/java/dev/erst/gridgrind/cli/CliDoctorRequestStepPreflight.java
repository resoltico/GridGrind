package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.json.DuplicateStepId;
import dev.erst.gridgrind.contract.json.ExplicitNullField;
import dev.erst.gridgrind.contract.json.FieldValidationProblem;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.contract.json.MessageInvariant;
import dev.erst.gridgrind.contract.json.MessageShape;
import dev.erst.gridgrind.contract.json.MissingRequiredField;
import dev.erst.gridgrind.contract.json.MissingTypeDiscriminator;
import dev.erst.gridgrind.contract.json.NonXlsxPath;
import dev.erst.gridgrind.contract.json.UnknownField;
import dev.erst.gridgrind.contract.json.UnknownTypeValue;
import dev.erst.gridgrind.contract.json.UnsupportedValue;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import tools.jackson.databind.node.ObjectNode;

/** Batches independently malformed step payloads before full doctor decoding begins. */
final class CliDoctorRequestStepPreflight {
  private CliDoctorRequestStepPreflight() {}

  static StepPreflight from(
      ObjectNode root, ProblemContextRequestSurfaces.RequestInput requestInput) {
    Objects.requireNonNull(root, "root must not be null");
    Objects.requireNonNull(requestInput, "requestInput must not be null");
    var stepsNode = root.get("steps");
    if (!(stepsNode instanceof tools.jackson.databind.node.ArrayNode stepsArray)) {
      return new StepPreflight(List.of(), false, true);
    }

    List<GridGrindProblemDetail.Problem> problems = new ArrayList<>();
    boolean usesSyntheticValues = false;
    boolean summaryTrustworthy = true;
    String executionModeType =
        CliDoctorRequestPlaceholderStepFactory.authoredExecutionModeType(root);
    for (int index = 0; index < stepsArray.size(); index++) {
      var authoredStep = stepsArray.get(index);
      try {
        decodeSingleStep(authoredStep);
      } catch (InvalidRequestException failure) {
        problems.add(rebasedStepProblem(requestInput, index, failure));
        stepsArray.set(
            index,
            CliDoctorRequestPlaceholderStepFactory.placeholderStepNode(
                authoredStep, index, executionModeType));
        usesSyntheticValues = true;
        summaryTrustworthy = false;
      } catch (InvalidRequestShapeException failure) {
        problems.add(rebasedStepProblem(requestInput, index, failure));
        stepsArray.set(
            index,
            CliDoctorRequestPlaceholderStepFactory.placeholderStepNode(
                authoredStep, index, executionModeType));
        usesSyntheticValues = true;
        summaryTrustworthy = false;
      }
    }
    return new StepPreflight(problems, usesSyntheticValues, summaryTrustworthy);
  }

  record StepPreflight(
      List<GridGrindProblemDetail.Problem> problems,
      boolean usesSyntheticValues,
      boolean summaryTrustworthy) {
    StepPreflight {
      problems = List.copyOf(Objects.requireNonNull(problems, "problems must not be null"));
    }
  }

  private static void decodeSingleStep(tools.jackson.databind.JsonNode stepNode) {
    ObjectNode request =
        GridGrindJsonOutput.requestTree(GridGrindProtocolCatalog.requestTemplate());
    request.withArray("steps").add(stepNode.deepCopy());
    GridGrindJson.readRequest(request.toString());
  }

  private static GridGrindProblemDetail.Problem rebasedStepProblem(
      ProblemContextRequestSurfaces.RequestInput requestInput,
      int stepIndex,
      InvalidRequestException failure) {
    InvalidRequestException rebased = rebaseStepFailure(failure, stepIndex);
    return GridGrindProblems.fromException(
        rebased,
        new ProblemContext.ReadRequest(
            requestInput,
            rebased
                .jsonPath()
                .map(ProblemContextRequestSurfaces.JsonLocation::pathOnly)
                .orElseGet(ProblemContextRequestSurfaces.JsonLocation::unavailable)));
  }

  private static GridGrindProblemDetail.Problem rebasedStepProblem(
      ProblemContextRequestSurfaces.RequestInput requestInput,
      int stepIndex,
      InvalidRequestShapeException failure) {
    InvalidRequestShapeException rebased = rebaseStepFailure(failure, stepIndex);
    return GridGrindProblems.fromException(
        rebased,
        new ProblemContext.ReadRequest(
            requestInput,
            rebased
                .jsonPath()
                .map(ProblemContextRequestSurfaces.JsonLocation::pathOnly)
                .orElseGet(ProblemContextRequestSurfaces.JsonLocation::unavailable)));
  }

  private static InvalidRequestException rebaseStepFailure(
      InvalidRequestException failure, int stepIndex) {
    dev.erst.gridgrind.contract.json.RequestProblemDescriptor.Invariant rebasedProblem =
        (dev.erst.gridgrind.contract.json.RequestProblemDescriptor.Invariant)
            rebaseStepProblem(failure.requestProblem(), stepIndex);
    return new InvalidRequestException(
        rebasedProblem,
        rebasedProblem.jsonPath(),
        Optional.empty(),
        Optional.empty(),
        failure.getCause());
  }

  private static InvalidRequestShapeException rebaseStepFailure(
      InvalidRequestShapeException failure, int stepIndex) {
    dev.erst.gridgrind.contract.json.RequestProblemDescriptor.Shape rebasedProblem =
        (dev.erst.gridgrind.contract.json.RequestProblemDescriptor.Shape)
            rebaseStepProblem(failure.requestProblem(), stepIndex);
    return new InvalidRequestShapeException(
        rebasedProblem,
        rebasedProblem.jsonPath(),
        Optional.empty(),
        Optional.empty(),
        failure.getCause());
  }

  static dev.erst.gridgrind.contract.json.RequestProblemDescriptor rebaseStepProblem(
      dev.erst.gridgrind.contract.json.RequestProblemDescriptor requestProblem, int stepIndex) {
    Optional<String> rebasedJsonPath =
        requestProblem.jsonPath().map(path -> rebaseStepPath(path, stepIndex));
    return switch (requestProblem) {
      case MissingRequiredField missingRequiredField ->
          new MissingRequiredField(rebaseStepPath(missingRequiredField.jsonPathValue(), stepIndex));
      case MissingTypeDiscriminator missingTypeDiscriminator ->
          new MissingTypeDiscriminator(
              rebaseStepPath(missingTypeDiscriminator.jsonPathValue(), stepIndex));
      case UnknownField unknownField ->
          new UnknownField(rebaseStepPath(unknownField.jsonPathValue(), stepIndex));
      case UnknownTypeValue unknownTypeValue ->
          new UnknownTypeValue(
              unknownTypeValue.typeId(),
              unknownTypeValue.jsonPath().map(path -> rebaseStepPath(path, stepIndex)),
              unknownTypeValue.similarValues(),
              unknownTypeValue.specificGuidance());
      case UnsupportedValue unsupportedValue ->
          new UnsupportedValue(
              unsupportedValue.value(),
              unsupportedValue.jsonPath().map(path -> rebaseStepPath(path, stepIndex)),
              unsupportedValue.allowedValues());
      case ExplicitNullField explicitNullField ->
          new ExplicitNullField(rebaseStepPath(explicitNullField.jsonPathValue(), stepIndex));
      case dev.erst.gridgrind.contract.json.ActionableShapeMessage actionableShapeMessage ->
          new dev.erst.gridgrind.contract.json.ActionableShapeMessage(
              actionableShapeMessage.message(),
              actionableShapeMessage.resolutionValue(),
              rebasedJsonPath);
      case MessageShape messageShape -> new MessageShape(messageShape.message(), rebasedJsonPath);
      case DuplicateStepId duplicateStepId ->
          new DuplicateStepId(
              duplicateStepId.value(), rebaseStepPath(duplicateStepId.jsonPathValue(), stepIndex));
      case NonXlsxPath nonXlsxPath ->
          new NonXlsxPath(
              nonXlsxPath.actualExtension(),
              nonXlsxPath.jsonPath().map(path -> rebaseStepPath(path, stepIndex)));
      case FieldValidationProblem fieldValidationProblem ->
          new FieldValidationProblem(
              fieldValidationProblem.fieldName(),
              rebasedJsonPath,
              fieldValidationProblem.rule(),
              fieldValidationProblem.operands());
      case dev.erst.gridgrind.contract.json.ActionableInvariantMessage actionableInvariantMessage ->
          new dev.erst.gridgrind.contract.json.ActionableInvariantMessage(
              actionableInvariantMessage.message(),
              actionableInvariantMessage.resolutionValue(),
              rebasedJsonPath);
      case MessageInvariant messageInvariant ->
          new MessageInvariant(messageInvariant.message(), rebasedJsonPath);
    };
  }

  private static String rebaseStepPath(String jsonPath, int stepIndex) {
    return jsonPath.startsWith("steps[0]")
        ? "steps[" + stepIndex + "]" + jsonPath.substring("steps[0]".length())
        : jsonPath;
  }
}
