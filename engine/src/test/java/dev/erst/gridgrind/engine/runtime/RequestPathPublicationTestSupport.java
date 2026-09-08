package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence.PublicationOutcome;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.nio.file.StandardOpenOption;
import java.util.stream.Stream;

/** Shares only publication-fixture behavior between the create-new and replacement test suites. */
final class RequestPathPublicationTestSupport {
  private RequestPathPublicationTestSupport() {}

  static PublicationOutcome.StagedArtifactVerification verifiedStage() {
    return new PublicationOutcome.StagedArtifactVerification();
  }

  static void writeOneByteThenFail(
      Path stagedArtifact, RequestPathBinding binding, Path siblingName) throws IOException {
    try (var sibling =
        binding.openSiblingChannel(
            siblingName,
            java.util.Set.of(
                StandardOpenOption.CREATE_NEW,
                StandardOpenOption.WRITE,
                LinkOption.NOFOLLOW_LINKS))) {
      sibling.write(ByteBuffer.wrap(Files.readAllBytes(stagedArtifact), 0, 1));
    }
    throw new IOException("synthetic private sibling short write");
  }

  static void copySiblingExactly(Path stagedArtifact, RequestPathBinding binding, Path siblingName)
      throws IOException {
    RequestPathPublicationFileSupport.copyAndForce(
        stagedArtifact,
        binding.openSiblingChannel(
            siblingName,
            java.util.Set.of(
                StandardOpenOption.CREATE_NEW,
                StandardOpenOption.WRITE,
                LinkOption.NOFOLLOW_LINKS)));
  }

  static void writeDigestMismatchSibling(
      Path stagedArtifact, RequestPathBinding binding, Path siblingName) throws IOException {
    assertTrue(Files.isRegularFile(stagedArtifact));
    try (var sibling =
        binding.openSiblingChannel(
            siblingName,
            java.util.Set.of(
                StandardOpenOption.CREATE_NEW,
                StandardOpenOption.WRITE,
                LinkOption.NOFOLLOW_LINKS))) {
      sibling.write(ByteBuffer.wrap(new byte[] {0, 0, 0}));
    }
  }

  static void moveAtomically(
      Path source, Path destination, WorkbookArtifactWriteDisposition disposition)
      throws IOException {
    switch (disposition) {
      case CREATE_NEW -> Files.move(source, destination, StandardCopyOption.ATOMIC_MOVE);
      case REPLACE_EXISTING ->
          Files.move(
              source,
              destination,
              StandardCopyOption.ATOMIC_MOVE,
              StandardCopyOption.REPLACE_EXISTING);
    }
  }

  static void assertNoPrivatePublicationSibling(Path root) throws IOException {
    try (Stream<Path> paths = Files.list(root)) {
      assertTrue(
          paths.noneMatch(path -> path.getFileName().toString().startsWith(".gridgrind-publish-")));
    }
  }
}
