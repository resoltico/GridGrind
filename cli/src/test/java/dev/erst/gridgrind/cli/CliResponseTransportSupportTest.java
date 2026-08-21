package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.io.IOException;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.atomic.AtomicReference;
import org.junit.jupiter.api.Test;

/** Tests response-file publication semantics. */
class CliResponseTransportSupportTest {
  @Test
  void successfulStagedPublicationExposesOnlyTheCompletePayload() throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-response-transport-success-");
    Path target = directory.resolve("response.json");
    byte[] payload = "complete payload".getBytes(java.nio.charset.StandardCharsets.UTF_8);

    CliResponseTransportSupport.writePayload(target, payload);

    assertArrayEquals(
        "complete payload\n".getBytes(java.nio.charset.StandardCharsets.UTF_8),
        Files.readAllBytes(target));
    assertEquals(List.of("response.json"), childNames(directory));
  }

  @Test
  void stagingFailureLeavesNoPartialResponseOrScratchFile() throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-response-transport-failure-");
    Path target = directory.resolve("response.json");

    IOException failure =
        assertThrows(
            IOException.class,
            () ->
                CliResponseTransportSupport.writePayload(
                    target,
                    "complete payload".getBytes(java.nio.charset.StandardCharsets.UTF_8),
                    (temporaryPath, payload) -> {
                      Files.write(temporaryPath, new byte[] {1, 2, 3});
                      throw new IOException("simulated staging failure");
                    }));

    assertEquals("simulated staging failure", failure.getMessage());
    assertFalse(Files.exists(target));
    assertEquals(List.of(), childNames(directory));
  }

  @Test
  void stagingCleanupFailureIsSuppressedOnTheOriginalPublicationFailure() throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-response-transport-cleanup-");
    Path target = directory.resolve("response.json");
    AtomicReference<Path> stagingPath = new AtomicReference<>();

    try {
      IOException failure =
          assertThrows(
              IOException.class,
              () ->
                  CliResponseTransportSupport.writePayload(
                      target,
                      "complete payload".getBytes(java.nio.charset.StandardCharsets.UTF_8),
                      (temporaryPath, payload) -> {
                        stagingPath.set(temporaryPath);
                        Files.delete(temporaryPath);
                        Files.createDirectory(temporaryPath);
                        Files.writeString(temporaryPath.resolve("retained"), "test artifact");
                        throw new IOException("simulated staging failure");
                      }));

      assertEquals("simulated staging failure", failure.getMessage());
      assertEquals(1, failure.getSuppressed().length);
      assertEquals(
          java.nio.file.DirectoryNotEmptyException.class, failure.getSuppressed()[0].getClass());
      assertFalse(Files.exists(target));
      assertEquals(List.of(stagingPath.get().getFileName().toString()), childNames(directory));
    } finally {
      Path temporaryPath = stagingPath.get();
      if (temporaryPath != null) {
        Files.deleteIfExists(temporaryPath.resolve("retained"));
        Files.deleteIfExists(temporaryPath);
      }
    }
  }

  @Test
  void existingResponseFileIsUntouchedWhenPublicationIsRejected() throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-response-transport-existing-");
    Path target = directory.resolve("response.json");
    byte[] existing = "existing response\n".getBytes(java.nio.charset.StandardCharsets.UTF_8);
    Files.write(target, existing);

    assertThrows(
        FileAlreadyExistsException.class,
        () ->
            CliResponseTransportSupport.writePayload(
                target, "replacement".getBytes(java.nio.charset.StandardCharsets.UTF_8)));

    assertArrayEquals(existing, Files.readAllBytes(target));
    assertEquals(List.of("response.json"), childNames(directory));
  }

  private static List<String> childNames(Path directory) throws IOException {
    try (var children = Files.list(directory)) {
      return children.map(path -> path.getFileName().toString()).sorted().toList();
    }
  }
}
