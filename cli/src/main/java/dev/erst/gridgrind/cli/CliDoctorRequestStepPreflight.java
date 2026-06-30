package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.ObjectNode;

/** Batches independently malformed step payloads before full doctor decoding begins. */
final class CliDoctorRequestStepPreflight {
  private CliDoctorRequestStepPreflight() {}

  static StepPreflight from(
      ObjectNode root, ProblemContextRequestSurfaces.RequestInput requestInput) {
    Objects.requireNonNull(root, "root must not be null");
    Objects.requireNonNull(requestInput, "requestInput must not be null");
    JsonNode stepsNode = root.get("steps");
    if (!(stepsNode instanceof ArrayNode stepsArray)) {
      return new StepPreflight(List.of(), false, true);
    }

    List<GridGrindProblemDetail.Problem> problems = new ArrayList<>();
    boolean usesSyntheticValues = false;
    boolean summaryTrustworthy = true;
    String executionModeType =
        CliDoctorRequestPlaceholderStepFactory.authoredExecutionModeType(root);
    for (int index = 0; index < stepsArray.size(); index++) {
      JsonNode authoredStep = stepsArray.get(index);
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

  private static void decodeSingleStep(JsonNode stepNode) {
    ObjectNode request = GridGrindJson.requestTree(GridGrindProtocolCatalog.requestTemplate());
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
    Optional<String> rebasedPath = failure.jsonPath().map(path -> rebaseStepPath(path, stepIndex));
    String rebasedMessage =
        rebaseMessagePath(Objects.requireNonNullElse(failure.getMessage(), ""), stepIndex);
    return new InvalidRequestException(
        rebasedMessage, rebasedPath, Optional.empty(), Optional.empty(), failure.getCause());
  }

  private static InvalidRequestShapeException rebaseStepFailure(
      InvalidRequestShapeException failure, int stepIndex) {
    Optional<String> rebasedPath = failure.jsonPath().map(path -> rebaseStepPath(path, stepIndex));
    String rebasedMessage =
        rebaseMessagePath(Objects.requireNonNullElse(failure.getMessage(), ""), stepIndex);
    return new InvalidRequestShapeException(
        rebasedMessage, rebasedPath, Optional.empty(), Optional.empty(), failure.getCause());
  }

  static String rebaseMessagePath(String message, int stepIndex) {
    Optional<String> messagePath = GridGrindRequestProblemSupport.jsonPathFromMessage(message);
    if (messagePath.isEmpty()) {
      return message;
    }
    String originalPath = messagePath.orElseThrow();
    String rebasedPath = rebaseStepPath(originalPath, stepIndex);
    if (originalPath.equals(rebasedPath)) {
      return message;
    }
    return message.replace("'" + originalPath + "'", "'" + rebasedPath + "'");
  }

  private static String rebaseStepPath(String jsonPath, int stepIndex) {
    return jsonPath.startsWith("steps[0]")
        ? "steps[" + stepIndex + "]" + jsonPath.substring("steps[0]".length())
        : jsonPath;
  }
}
