package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliPrimaryOutputException;
import dev.erst.gridgrind.cli.io.CliResponseWriter;
import dev.erst.gridgrind.core.result.CommandError;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Last-resort structured failure emission for unexpected CLI/runtime faults. */
public final class CliUnexpectedFailureSupport {
  private CliUnexpectedFailureSupport() {}

  /** Emits a sanitized failure through the selected primary transport. */
  public static int emit(
      String[] args,
      Optional<Path> responsePathHint,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      Throwable exception) {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(responsePathHint, "responsePathHint must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(exception, "exception must not be null");

    CommandError diagnostic = CommandErrors.unexpectedFailure(CommandName.primary(args), exception);
    try {
      return new CliResponseWriter()
          .writeCommandError(responsePathHint, stdout, stderr, diagnostic, prettyJson);
    } catch (IOException | CliPrimaryOutputException ignored) {
      return CliExitCodes.forCommandError(diagnostic);
    }
  }
}
