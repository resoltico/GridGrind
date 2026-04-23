package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.catalog.GridGrindShippedExamples;
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

/** Writes the checkout-rooted example request fixtures from the authoritative contract registry. */
public final class ExampleRequestFixturesWriter {
  private static final byte[] LINE_SEPARATOR_BYTES =
      System.lineSeparator().getBytes(StandardCharsets.UTF_8);

  private ExampleRequestFixturesWriter() {}

  public static void main(String[] args) throws IOException {
    Path examplesDirectory = examplesDirectory();
    List<GridGrindShippedExamples.ShippedExample> examples =
        GridGrindShippedExamples.repositoryExamples();
    Set<String> expectedFiles =
        examples.stream()
            .map(GridGrindShippedExamples.ShippedExample::fileName)
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

    for (GridGrindShippedExamples.ShippedExample example : examples) {
      Files.write(
          examplesDirectory.resolve(example.fileName()), requestBytesWithLineSeparator(example));
    }
  }

  private static byte[] requestBytesWithLineSeparator(
      GridGrindShippedExamples.ShippedExample example) throws IOException {
    byte[] requestBytes = GridGrindJson.writeRequestBytes(example.plan());
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
