package dev.erst.gridgrind.cli.golden;

import java.io.IOException;
import java.nio.file.Path;

/** Runs a host invocation while retaining its exit code and both output streams. */
@FunctionalInterface
interface WorkbookLauncher {
  GoldenLauncher.Result run(Path workspace, String... arguments)
      throws IOException, InterruptedException;
}
