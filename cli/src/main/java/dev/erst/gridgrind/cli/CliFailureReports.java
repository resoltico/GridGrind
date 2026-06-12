package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliFailureLocation;
import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Shared constructors for machine-readable CLI failure reports. */
final class CliFailureReports {
  private CliFailureReports() {}

  static CliFailureReport invalidArguments(
      int exitCode,
      String command,
      String phase,
      Optional<String> argument,
      String message,
      List<String> suggestions,
      Optional<String> resolution) {
    return report(
        exitCode,
        command,
        phase,
        GridGrindProblemCode.INVALID_ARGUMENTS,
        Optional.empty(),
        argument,
        message,
        suggestions,
        resolution);
  }

  static CliFailureReport report(
      int exitCode,
      String command,
      String phase,
      GridGrindProblemCode code,
      Optional<CliFailureLocation> location,
      Optional<String> argument,
      String message,
      List<String> suggestions,
      Optional<String> resolution) {
    Objects.requireNonNull(code, "code must not be null");
    return new CliFailureReport(
        GridGrindProtocolVersion.current(),
        exitCode,
        command,
        phase,
        code,
        message,
        Objects.requireNonNull(location, "location must not be null"),
        Objects.requireNonNull(argument, "argument must not be null"),
        List.copyOf(Objects.requireNonNull(suggestions, "suggestions must not be null")),
        Objects.requireNonNull(resolution, "resolution must not be null"));
  }

  static CliFailureReport responseWriteFailure(
      String command,
      String payloadName,
      Path targetPath,
      java.io.IOException exception,
      Optional<String> stdoutSuggestion) {
    Objects.requireNonNull(targetPath, "targetPath must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    Objects.requireNonNull(stdoutSuggestion, "stdoutSuggestion must not be null");
    return report(
        1,
        command,
        "write-response",
        GridGrindProblemCode.IO_ERROR,
        Optional.empty(),
        Optional.of("--response"),
        CliResponseWriter.responseWriteMessage(exception, targetPath),
        stdoutSuggestion.filter(text -> !text.isBlank()).map(List::of).orElse(List.of()),
        Optional.of(
            "Provide one writable file path after --response, or remove --response to print the "
                + payloadName
                + " on stdout."));
  }

  static CliFailureReport readRequestFailure(
      int exitCode,
      String command,
      Optional<String> argument,
      GridGrindProblemDetail.Problem problem,
      Throwable exception) {
    Objects.requireNonNull(exception, "exception must not be null");
    Objects.requireNonNull(problem, "problem must not be null");
    return new CliFailureReport(
        GridGrindProtocolVersion.current(),
        exitCode,
        command,
        "read-request",
        problem.code(),
        problem.message(),
        locationForReadRequest(problem, exception),
        Objects.requireNonNull(argument, "argument must not be null"),
        suggestionsForReadRequest(problem.code(), command),
        resolutionForReadRequest(problem, command));
  }

  private static List<String> suggestionsForReadRequest(GridGrindProblemCode code, String command) {
    return switch (code) {
      case INVALID_JSON ->
          List.of(
              "gridgrind --print-request-template --response request.json",
              "gridgrind --help-protocol");
      case INVALID_REQUEST_SHAPE ->
          "doctor-request".equals(command)
              ? List.of(
                  "gridgrind --print-protocol-catalog --search \"sheet layout\"",
                  "gridgrind --print-request-template --response request.json")
              : List.of(
                  "gridgrind --print-protocol-catalog --search \"sheet layout\"",
                  "gridgrind --doctor-request --request request.json",
                  "gridgrind --help-protocol");
      case INVALID_REQUEST ->
          "doctor-request".equals(command)
              ? List.of("gridgrind --print-request-template --response request.json")
              : List.of(
                  "gridgrind --doctor-request --request request.json", "gridgrind --help-protocol");
      case IO_ERROR -> List.of("gridgrind --request request.json", "gridgrind --help-guidance");
      default -> List.of("gridgrind --help", "gridgrind --help-protocol");
    };
  }

  private static Optional<String> resolutionForReadRequest(
      GridGrindProblemDetail.Problem problem, String command) {
    GridGrindProblemCode code = problem.code();
    Optional<String> specificResolution =
        GridGrindRequestProblemSupport.specificResolution(
            code, problem.message(), problem.context());
    return switch (code) {
      case INVALID_JSON ->
          Optional.of(
              "Provide one complete JSON request document. Start from --print-request-template"
                  + " when you need the canonical shape.");
      case INVALID_REQUEST_SHAPE ->
          specificResolution
              .map(
                  resolution ->
                      resolution
                          + " Use --print-protocol-catalog --search \"sheet layout\" or"
                          + " --help-protocol when you need the authoritative field and"
                          + " discriminator contract.")
              .or(
                  () ->
                      Optional.of(
                          "The JSON document parsed, but its field layout does not match the"
                              + " request contract. Use --print-protocol-catalog --search"
                              + " \"sheet layout\" to discover valid type discriminator values,"
                              + " or compare against --help-protocol or a printed starter"
                              + " request."));
      case INVALID_REQUEST ->
          specificResolution
              .map(
                  resolution ->
                      resolution
                          + " "
                          + ("doctor-request".equals(command)
                              ? "Rerun --doctor-request after correcting the request."
                              : "Run --doctor-request first, then execute the corrected request."))
              .or(
                  () ->
                      Optional.of(
                          "The request document decoded but violates request invariants. "
                              + ("doctor-request".equals(command)
                                  ? "Correct the authored request and rerun --doctor-request."
                                  : "Run --doctor-request first, then execute the corrected"
                                      + " request.")));
      case IO_ERROR ->
          Optional.of(
              "Make sure --request points at one readable JSON file and that the process can"
                  + " traverse its parent directories.");
      default -> Optional.of("Correct the request input, then rerun the command.");
    };
  }

  private static Optional<CliFailureLocation> locationForReadRequest(
      GridGrindProblemDetail.Problem problem, Throwable exception) {
    Optional<CliFailureLocation> exceptionLocation = CliFailureLocation.from(exception);
    if (exceptionLocation.isPresent()) {
      return exceptionLocation;
    }
    if (problem.context() instanceof ProblemContext.ReadRequest readRequest) {
      return CliFailureLocation.from(
          readRequest.jsonPath(), readRequest.jsonLine(), readRequest.jsonColumn());
    }
    return Optional.empty();
  }

  static CliFailureReport unexpectedFailure(String command, String phase, Throwable exception) {
    Objects.requireNonNull(command, "command must not be null");
    Objects.requireNonNull(phase, "phase must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    String message =
        Optional.ofNullable(exception.getMessage())
            .filter(text -> !text.isBlank())
            .orElse(GridGrindProblemCode.INTERNAL_ERROR.title());
    return report(
        1,
        command,
        phase,
        GridGrindProblemCode.INTERNAL_ERROR,
        Optional.empty(),
        Optional.empty(),
        message,
        List.of("gridgrind --help", "gridgrind --help-protocol"),
        Optional.of(GridGrindProblemCode.INTERNAL_ERROR.resolution()));
  }
}
