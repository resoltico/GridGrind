package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.core.problem.RequestDoctorReport;
import dev.erst.gridgrind.core.result.CommandError;
import dev.erst.gridgrind.core.result.WorkbookResult;
import java.util.Objects;

/** Maps already-classified command outcomes to the CLI process contract. */
public final class CliExitCodes {
  private CliExitCodes() {}

  /** Returns zero for successful workbook execution and one for execution failure. */
  public static int forWorkbookResult(WorkbookResult result) {
    return switch (Objects.requireNonNull(result, "result must not be null")) {
      case WorkbookResult.Success _ -> 0;
      case WorkbookResult.Failure _ -> 1;
    };
  }

  /** Returns two for argument or request rejection and one for operational failure. */
  public static int forCommandError(CommandError commandError) {
    Objects.requireNonNull(commandError, "commandError must not be null");
    return switch (commandError.primaryProblem().category()) {
      case ARGUMENTS, REQUEST -> 2;
      default -> 1;
    };
  }

  /** Returns zero only when diagnosis reports a valid request. */
  public static int forDoctorReport(RequestDoctorReport report) {
    return Objects.requireNonNull(report, "report must not be null").valid() ? 0 : 1;
  }
}
