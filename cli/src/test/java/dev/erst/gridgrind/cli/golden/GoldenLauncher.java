package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.GridGrindCli;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

/** Exercises the public transport both inside the JVM and through the delivered artifact. */
enum GoldenLauncher {
  IN_PROCESS,
  SHADOW_JAR;

  static final Path REPOSITORY = Path.of(System.getProperty("gridgrind.repository"));

  Result run(Path workspace, String... arguments) throws IOException, InterruptedException {
    if (this == IN_PROCESS) {
      ByteArrayOutputStream stdout = new ByteArrayOutputStream();
      ByteArrayOutputStream stderr = new ByteArrayOutputStream();
      int exit =
          new GridGrindCli().run(arguments, new ByteArrayInputStream(new byte[0]), stdout, stderr);
      return new Result(
          exit, stdout.toString(StandardCharsets.UTF_8), stderr.toString(StandardCharsets.UTF_8));
    }
    List<String> command = new ArrayList<>();
    command.add(Path.of(System.getProperty("java.home"), "bin", "java").toString());
    command.add("-jar");
    command.add(REPOSITORY.resolve("cli/build/libs/gridgrind.jar").toString());
    command.addAll(List.of(arguments));
    Path stdout = workspace.resolve("process-stdout.log");
    Path stderr = workspace.resolve("process-stderr.log");
    try (Process process =
        new ProcessBuilder(command)
            .directory(workspace.toFile())
            .redirectOutput(stdout.toFile())
            .redirectError(stderr.toFile())
            .start()) {
      try {
        process.getOutputStream().close();
        assertTrue(process.waitFor(60, TimeUnit.SECONDS), () -> "CLI timed out: " + command);
        return new Result(process.exitValue(), Files.readString(stdout), Files.readString(stderr));
      } finally {
        if (process.isAlive()) {
          process.destroyForcibly();
          process.waitFor();
        }
      }
    } finally {
      Files.deleteIfExists(stdout);
      Files.deleteIfExists(stderr);
    }
  }

  record Result(int exit, String stdout, String stderr) {}
}
