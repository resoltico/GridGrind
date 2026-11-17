package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliResponseWriter;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.Optional;

/** One CLI mapping for authority intake failures across run/doctor and primary transports. */
final class CliAuthorityFailureSupport {
  private CliAuthorityFailureSupport() {}

  static int write(
      String[] args,
      Optional<Path> responsePath,
      boolean pretty,
      OutputStream stdout,
      OutputStream stderr,
      Path authorityPath,
      Throwable failure)
      throws IOException {
    var error = commandError(CommandName.primary(args), authorityPath, failure);
    return new CliResponseWriter().writeCommandError(responsePath, stdout, stderr, error, pretty);
  }

  static dev.erst.gridgrind.core.result.CommandError commandError(
      String command, Path authorityPath, Throwable failure) {
    var context =
        new ProblemContext.ReadAuthority(authorityPath.toString(), JsonLocation.unavailable());
    return CommandErrors.commandError(command, GridGrindProblems.fromException(failure, context));
  }
}
