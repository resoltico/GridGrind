package dev.erst.gridgrind.fuzz;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

/** Replays every committed input through the same checks used by the standalone driver. */
class SeedReplayTest {
  @ParameterizedTest(name = "{0}/{1}")
  @MethodSource("seeds")
  void committedInput(String harness, String name, Path input) throws IOException {
    Harness.fromKey(harness).check(Files.readAllBytes(input));
  }

  static Stream<Arguments> seeds() throws IOException {
    Path root = Path.of(System.getProperty("gridgrind.fuzz.seeds"));
    List<Arguments> cases = new ArrayList<>();
    for (Harness harness : Harness.values()) {
      try (Stream<Path> files = Files.list(root.resolve(harness.key()))) {
        List<Path> inputs = files.filter(Files::isRegularFile).sorted().toList();
        if (inputs.isEmpty())
          throw new IllegalStateException("No committed seeds for " + harness.key());
        for (Path input : inputs)
          cases.add(Arguments.of(harness.key(), input.getFileName().toString(), input));
      }
    }
    return cases.stream();
  }
}
