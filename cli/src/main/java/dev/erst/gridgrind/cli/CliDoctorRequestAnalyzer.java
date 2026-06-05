package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import dev.erst.gridgrind.engine.api.GridGrindRequestRequirements;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.ObjectNode;

/** Builds machine-readable doctor reports from raw request payloads before execution begins. */
final class CliDoctorRequestAnalyzer {
  private static final String SYNTHETIC_EXISTING_SOURCE_PATH = "__gridgrind_missing_source__.xlsx";
  private static final String SYNTHETIC_SAVE_AS_PATH = "__gridgrind_missing_output__.xlsx";

  private final GridGrindRequestDoctor requestDoctor;

  CliDoctorRequestAnalyzer(GridGrindRequestDoctor requestDoctor) {
    this.requestDoctor = GridGrindRequestDoctor.requireNonNull(requestDoctor);
  }

  RequestDoctorReport diagnose(
      Optional<Path> requestPath,
      Optional<Path> executionRootPath,
      Optional<Path> tempRootPath,
      byte[] requestBytes,
      InputStream stdin)
      throws IOException {
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(executionRootPath, "executionRootPath must not be null");
    Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
    Objects.requireNonNull(requestBytes, "requestBytes must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");

    ProblemContextRequestSurfaces.RequestInput requestInput = requestInput(requestPath);
    JsonNode requestTree;
    try {
      requestTree = GridGrindJson.readRequestTree(requestBytes);
    } catch (InvalidJsonException exception) {
      return RequestDoctorReport.invalid(
          Optional.empty(),
          List.of(),
          List.of(GridGrindProblems.fromException(exception, readRequestContext(requestInput))));
    }

    TreePreflight preflight = TreePreflight.from(requestTree, requestInput);
    WorkbookPlan request;
    try {
      request =
          GridGrindJson.readRequest(
              preflight.sanitizedTree().toString().getBytes(StandardCharsets.UTF_8));
    } catch (InvalidRequestException | InvalidRequestShapeException exception) {
      return RequestDoctorReport.invalid(
          Optional.empty(),
          List.of(),
          mergeProblems(
              preflight.problems(),
              List.of(
                  GridGrindProblems.fromException(exception, readRequestContext(requestInput)))));
    }

    RequestDoctorReport baseReport =
        runBaseDoctorReport(
            requestPath,
            executionRootPath,
            tempRootPath,
            stdin,
            preflight.usesSyntheticValues(),
            request);
    if (requestPath.isEmpty() && GridGrindRequestRequirements.requiresStandardInput(request)) {
      GridGrindProblemDetail.Problem standardInputProblem =
          GridGrindProblems.problem(
              GridGrindProblemCode.INVALID_REQUEST,
              GridGrindContractText.standardInputRequiresRequestMessage(),
              new ProblemContext.ValidateRequest(requestShape(request)),
              List.of());
      return RequestDoctorReport.invalid(
          baseReport.summary(),
          baseReport.warnings(),
          mergeProblems(
              List.of(standardInputProblem),
              mergeProblems(preflight.problems(), mergeableProblems(baseReport))));
    }
    List<GridGrindProblemDetail.Problem> mergedProblems =
        mergeProblems(preflight.problems(), mergeableProblems(baseReport));
    if (!mergedProblems.isEmpty()) {
      return RequestDoctorReport.invalid(
          baseReport.summary(), baseReport.warnings(), mergedProblems);
    }
    return baseReport;
  }

  private RequestDoctorReport runBaseDoctorReport(
      Optional<Path> requestPath,
      Optional<Path> executionRootPath,
      Optional<Path> tempRootPath,
      InputStream stdin,
      boolean usesSyntheticValues,
      WorkbookPlan request)
      throws IOException {
    if (usesSyntheticValues) {
      return requestDoctor.diagnose(request);
    }
    if (requestPath.isPresent() || executionRootPath.isPresent()) {
      return requestDoctor.diagnose(
          request,
          CliExecutionBindingsFactory.create(
              requestPath, executionRootPath, tempRootPath, request, stdin));
    }
    return requestDoctor.diagnose(request);
  }

  private static List<GridGrindProblemDetail.Problem> mergeableProblems(
      RequestDoctorReport report) {
    Objects.requireNonNull(report, "report must not be null");
    return report.valid() ? List.of() : report.problems();
  }

  private static List<GridGrindProblemDetail.Problem> mergeProblems(
      List<GridGrindProblemDetail.Problem> left, List<GridGrindProblemDetail.Problem> right) {
    java.util.Set<String> seen = new java.util.HashSet<>(left.size() + right.size());
    List<GridGrindProblemDetail.Problem> merged = new ArrayList<>(left.size() + right.size());
    for (GridGrindProblemDetail.Problem problem : left) {
      seen.add(problemKey(problem));
      merged.add(problem);
    }
    for (GridGrindProblemDetail.Problem problem : right) {
      if (seen.add(problemKey(problem))) {
        merged.add(problem);
      }
    }
    return List.copyOf(merged);
  }

  private static String problemKey(GridGrindProblemDetail.Problem problem) {
    return problem.code() + "|" + problem.context().stage() + "|" + problem.message();
  }

  private static ProblemContext.ReadRequest readRequestContext(
      ProblemContextRequestSurfaces.RequestInput requestInput) {
    return new ProblemContext.ReadRequest(
        requestInput, ProblemContextRequestSurfaces.JsonLocation.unavailable());
  }

  private static ProblemContextRequestSurfaces.RequestInput requestInput(
      Optional<Path> requestPath) {
    return requestPath.isEmpty()
        ? ProblemContextRequestSurfaces.RequestInput.standardInput()
        : ProblemContextRequestSurfaces.RequestInput.requestFile(
            requestPath.orElseThrow().toAbsolutePath().toString());
  }

  private static ProblemContextRequestSurfaces.RequestShape requestShape(WorkbookPlan request) {
    return ProblemContextRequestSurfaces.RequestShape.known(
        switch (request.source()) {
          case WorkbookPlan.WorkbookSource.New _ -> "NEW";
          case WorkbookPlan.WorkbookSource.ExistingFile _ -> "EXISTING";
        },
        switch (request.persistence()) {
          case WorkbookPlan.WorkbookPersistence.None _ -> "NONE";
          case WorkbookPlan.WorkbookPersistence.OverwriteSource _ -> "OVERWRITE";
          case WorkbookPlan.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
        });
  }

  private record TreePreflight(
      ObjectNode sanitizedTree,
      List<GridGrindProblemDetail.Problem> problems,
      boolean usesSyntheticValues) {
    private TreePreflight {
      Objects.requireNonNull(sanitizedTree, "sanitizedTree must not be null");
      problems = List.copyOf(Objects.requireNonNull(problems, "problems must not be null"));
    }

    private static TreePreflight from(
        JsonNode requestTree, ProblemContextRequestSurfaces.RequestInput requestInput) {
      Objects.requireNonNull(requestTree, "requestTree must not be null");
      Objects.requireNonNull(requestInput, "requestInput must not be null");
      if (!(requestTree instanceof ObjectNode requestObject)) {
        return new TreePreflight(
            GridGrindJson.requestTree(GridGrindProtocolCatalog.requestTemplate()),
            List.of(
                GridGrindProblems.problem(
                    GridGrindProblemCode.INVALID_REQUEST_SHAPE,
                    "JSON request must be one object at the root",
                    new ProblemContext.ReadRequest(
                        requestInput, ProblemContextRequestSurfaces.JsonLocation.unavailable()),
                    List.of())),
            true);
      }

      ObjectNode sanitized = requestObject.deepCopy();
      List<GridGrindProblemDetail.Problem> problems = new ArrayList<>();
      boolean usesSyntheticValues =
          applyTemplateDefaults(
              sanitized,
              GridGrindJson.requestTree(GridGrindProtocolCatalog.requestTemplate()),
              "",
              requestInput,
              problems);
      usesSyntheticValues |=
          applyConditionalWorkbookPathDefaults(
              sanitized,
              requestInput,
              problems,
              "source",
              "EXISTING",
              "path",
              SYNTHETIC_EXISTING_SOURCE_PATH);
      usesSyntheticValues |=
          applyConditionalWorkbookPathDefaults(
              sanitized,
              requestInput,
              problems,
              "persistence",
              "SAVE_AS",
              "path",
              SYNTHETIC_SAVE_AS_PATH);
      return new TreePreflight(sanitized, problems, usesSyntheticValues);
    }

    private static boolean applyTemplateDefaults(
        ObjectNode authored,
        ObjectNode template,
        String path,
        ProblemContextRequestSurfaces.RequestInput requestInput,
        List<GridGrindProblemDetail.Problem> problems) {
      boolean mutated = false;
      for (var field : template.properties()) {
        String childPath = path.isEmpty() ? field.getKey() : path + "." + field.getKey();
        JsonNode existing = authored.get(field.getKey());
        if (existing == null || existing.isNull()) {
          authored.set(field.getKey(), field.getValue().deepCopy());
          problems.add(missingFieldProblem(requestInput, childPath));
          mutated = true;
          continue;
        }
        if (existing.isObject() && field.getValue().isObject()) {
          mutated |=
              applyTemplateDefaults(
                  (ObjectNode) existing,
                  (ObjectNode) field.getValue(),
                  childPath,
                  requestInput,
                  problems);
        }
      }
      return mutated;
    }

    private static boolean applyConditionalWorkbookPathDefaults(
        ObjectNode root,
        ProblemContextRequestSurfaces.RequestInput requestInput,
        List<GridGrindProblemDetail.Problem> problems,
        String objectField,
        String expectedType,
        String pathField,
        String syntheticPath) {
      JsonNode node = root.get(objectField);
      if (!(node instanceof ObjectNode objectNode)) {
        return false;
      }
      if (!hasExpectedWorkbookObjectType(objectNode, expectedType)) {
        return false;
      }
      if (hasNonBlankPath(objectNode.get(pathField))) {
        return false;
      }
      objectNode.put(pathField, syntheticPath);
      problems.add(missingFieldProblem(requestInput, objectField + "." + pathField));
      return true;
    }

    private static GridGrindProblemDetail.Problem missingFieldProblem(
        ProblemContextRequestSurfaces.RequestInput requestInput, String jsonPath) {
      return GridGrindProblems.fromException(
          new InvalidRequestException(
              "Missing required field '" + jsonPath + "'",
              Optional.of(jsonPath),
              Optional.empty(),
              Optional.empty(),
              null),
          new ProblemContext.ReadRequest(
              requestInput, ProblemContextRequestSurfaces.JsonLocation.pathOnly(jsonPath)));
    }

    private static boolean hasExpectedWorkbookObjectType(
        ObjectNode objectNode, String expectedType) {
      String authoredType = authoredScalarText(objectNode.path("type"));
      return expectedType.equals(authoredType);
    }

    private static boolean hasNonBlankPath(JsonNode pathNode) {
      return !authoredScalarText(pathNode).isBlank();
    }

    private static String authoredScalarText(JsonNode node) {
      if (node == null || node.isNull()) {
        return "";
      }
      if (node.isString()) {
        return node.stringValue();
      }
      if (node.isNumber()) {
        return node.numberValue().toString();
      }
      if (node.isBoolean()) {
        return Boolean.toString(node.booleanValue());
      }
      return "";
    }
  }
}
