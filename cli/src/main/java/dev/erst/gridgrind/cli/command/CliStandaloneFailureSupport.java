package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliResponseWriter;
import dev.erst.gridgrind.cli.logging.StandaloneLoggingRefusal;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Optional;

/** Maps a preinitialization logging-policy refusal through the existing primary transport. */
public final class CliStandaloneFailureSupport {
  private CliStandaloneFailureSupport() {}

  /** Reports only the qualified setting identity, never its supplied value. */
  public static int emit(
      String[] args, OutputStream stdout, OutputStream stderr, StandaloneLoggingRefusal refusal) {
    var error =
        CommandErrors.invalidArguments(
            CommandName.primary(args),
            Optional.of(refusal.setting()),
            "The standalone CLI requires its owned logging transport. Remove the conflicting"
                + " setting: "
                + refusal.setting());
    try {
      return new CliResponseWriter()
          .writeCommandError(
              CliPathArguments.responsePathHint(args),
              stdout,
              stderr,
              error,
              CliArguments.prettyJsonHint(args));
    } catch (IOException | dev.erst.gridgrind.cli.io.CliPrimaryOutputException failure) {
      return 1;
    }
  }
}
