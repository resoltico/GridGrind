package dev.erst.gridgrind.cli.io;

import dev.erst.gridgrind.core.discovery.ShippedExamples;
import dev.erst.gridgrind.json.codec.GridGrindJsonOutput;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/** Writes the checkout-rooted example request fixtures from the authoritative CLI registry. */
public final class ExampleRequestFixturesWriter {
  private static final byte[] LINE_SEPARATOR_BYTES =
      System.lineSeparator().getBytes(StandardCharsets.UTF_8);

  private ExampleRequestFixturesWriter() {}

  public static void main(String[] args) throws IOException {
    Path examplesDirectory = examplesDirectory();
    List<ShippedExamples.ShippedExample> examples = ShippedExamples.repositoryExamples();
    Set<String> expectedFiles =
        examples.stream()
            .map(ShippedExamples.ShippedExample::requestFileName)
            .collect(Collectors.toUnmodifiableSet());

    try (Stream<Path> stream = Files.list(examplesDirectory)) {
      for (Path path :
          stream
              .filter(candidate -> candidate.getFileName().toString().endsWith(".json"))
              .sorted(Comparator.comparing(Path::toString))
              .toList()) {
        if (!expectedFiles.contains(path.getFileName().toString())) {
          Files.delete(path);
        }
      }
    }

    for (ShippedExamples.ShippedExample example : examples) {
      Files.write(
          examplesDirectory.resolve(example.requestFileName()),
          requestBytesWithLineSeparator(example));
    }
  }

  private static byte[] requestBytesWithLineSeparator(ShippedExamples.ShippedExample example)
      throws IOException {
    byte[] requestBytes = GridGrindJsonOutput.writeRequestBytes(example.plan());
    byte[] output = Arrays.copyOf(requestBytes, requestBytes.length + LINE_SEPARATOR_BYTES.length);
    System.arraycopy(
        LINE_SEPARATOR_BYTES, 0, output, requestBytes.length, LINE_SEPARATOR_BYTES.length);
    return output;
  }

  private static Path examplesDirectory() {
    Path candidate = Path.of("").toAbsolutePath().normalize();
    while (candidate != null) {
      if (Files.exists(candidate.resolve("gradle.properties"))
          && Files.exists(candidate.resolve("examples"))) {
        return candidate.resolve("examples");
      }
      candidate = candidate.getParent();
    }
    throw new IllegalStateException("Could not locate the GridGrind examples directory.");
  }
}
