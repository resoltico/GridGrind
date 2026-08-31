package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CommandError;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import java.util.Objects;

/** Maps already-classified command outcomes to the CLI process contract. */
final class CliExitCodes {
  private CliExitCodes() {}

  static int forWorkbookResult(WorkbookResult result) {
    return switch (Objects.requireNonNull(result, "result must not be null")) {
      case WorkbookResult.Success _ -> 0;
      case WorkbookResult.Failure _ -> 1;
    };
  }

  static int forCommandError(CommandError commandError) {
    Objects.requireNonNull(commandError, "commandError must not be null");
    return switch (commandError.primaryProblem().category()) {
      case ARGUMENTS, REQUEST -> 2;
      default -> 1;
    };
  }

  static int forDoctorReport(RequestDoctorReport report) {
    return Objects.requireNonNull(report, "report must not be null").valid() ? 0 : 1;
  }
}
