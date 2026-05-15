package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliFailureLocation;
import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Shared constructors for machine-readable CLI failure reports. */
final class CliFailureReports {
  private CliFailureReports() {}

  static CliFailureReport invalidArguments(
      int exitCode,
      String command,
      Optional<String> argument,
      String message,
      List<String> suggestions,
      Optional<String> resolution) {
    return report(
        exitCode,
        command,
        GridGrindProblemCode.INVALID_ARGUMENTS,
        argument,
        message,
        suggestions,
        resolution);
  }

  static CliFailureReport report(
      int exitCode,
      String command,
      GridGrindProblemCode code,
      Optional<String> argument,
      String message,
      List<String> suggestions,
      Optional<String> resolution) {
    Objects.requireNonNull(code, "code must not be null");
    return new CliFailureReport(
        GridGrindProtocolVersion.current(),
        exitCode,
        command,
        code,
        message,
        CliFailureLocation.unavailable(),
        Objects.requireNonNull(argument, "argument must not be null"),
        List.copyOf(Objects.requireNonNull(suggestions, "suggestions must not be null")),
        Objects.requireNonNull(resolution, "resolution must not be null"));
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
        problem.code(),
        problem.message(),
        CliFailureLocation.from(exception),
        Objects.requireNonNull(argument, "argument must not be null"),
        suggestionsForReadRequest(problem.code(), command),
        resolutionForReadRequest(problem.code(), command));
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
                  "gridgrind --print-protocol-catalog --search <term>",
                  "gridgrind --print-request-template --response request.json")
              : List.of(
                  "gridgrind --print-protocol-catalog --search <term>",
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
      GridGrindProblemCode code, String command) {
    return switch (code) {
      case INVALID_JSON ->
          Optional.of(
              "Provide one complete JSON request document. Start from --print-request-template"
                  + " when you need the canonical shape.");
      case INVALID_REQUEST_SHAPE ->
          Optional.of(
              "The JSON document parsed, but its field layout does not match the request"
                  + " contract. Use --print-protocol-catalog --search <term> to discover"
                  + " valid type discriminator values, or compare against --help-protocol or"
                  + " a printed starter request.");
      case INVALID_REQUEST ->
          Optional.of(
              "The request document decoded but violates request invariants. "
                  + ("doctor-request".equals(command)
                      ? "Correct the authored request and rerun --doctor-request."
                      : "Run --doctor-request first, then execute the corrected request."));
      case IO_ERROR ->
          Optional.of(
              "Make sure --request points at one readable JSON file and that the process can"
                  + " traverse its parent directories.");
      default -> Optional.of("Correct the request input, then rerun the command.");
    };
  }
}
