package dev.erst.gridgrind.fuzz;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import edu.berkeley.cs.jqf.fuzz.guidance.Result;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.Random;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Independent raw-byte controls for seed identity, finite EOF and explicit size admission. */
class RawInputGuidanceTest {
  @TempDir Path directory;

  @Test
  void everyCommittedSeedReachesTheCursorWithoutFramingOrPadding() throws Exception {
    Path source = Path.of(System.getProperty("gridgrind.fuzz.seeds"));
    int count = 0;
    for (Harness harness : Harness.values()) {
      File[] seeds;
      try (var files = Files.list(source.resolve(harness.key()))) {
        seeds = files.filter(Files::isRegularFile).sorted().map(Path::toFile).toArray(File[]::new);
      }
      Path output = directory.resolve(harness.key());
      var guidance =
          new RawInputGuidance(
              harness.key(),
              Duration.ofMillis(1),
              seeds.length + 1L,
              output.toFile(),
              output.resolve("findings").toFile(),
              seeds,
              32_768,
              new Random(0));
      for (File seed : seeds) {
        assertTrue(guidance.hasInput(), "Initial seed admission must be explicit");
        try (var input = guidance.getInput()) {
          assertArrayEquals(
              Files.readAllBytes(seed.toPath()), input.readAllBytes(), seed.toString());
          assertEquals(-1, input.read());
        }
        guidance.handleResult(Result.SUCCESS, null);
        count++;
      }
    }
    assertTrue(count > 0, "The seed enumeration must not be vacuous");
  }

  @Test
  void emptyTapeHasFiniteRepeatableEof() throws Exception {
    Path seed = directory.resolve("empty");
    Files.write(seed, new byte[0]);
    var guidance =
        new RawInputGuidance(
            "empty",
            Duration.ofSeconds(1),
            2,
            directory.resolve("results").toFile(),
            directory.resolve("findings").toFile(),
            new File[] {seed.toFile()},
            32,
            new Random(0));
    try (var input = guidance.getInput()) {
      assertEquals(-1, input.read());
      assertEquals(-1, input.read());
    }
    guidance.handleResult(Result.SUCCESS, null);
  }

  @Test
  void configuredBoundCannotSilentlyTruncateASeed() throws Exception {
    Path seed = directory.resolve("too-large");
    Files.write(seed, new byte[] {1, 2, 3});
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RawInputGuidance(
                "bounds",
                Duration.ofSeconds(1),
                2,
                directory.resolve("results").toFile(),
                directory.resolve("findings").toFile(),
                new File[] {seed.toFile()},
                2,
                new Random(0)));
  }
}
