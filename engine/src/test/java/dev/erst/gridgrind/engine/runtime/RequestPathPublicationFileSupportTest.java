package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence.PublicationOutcome;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.channels.FileChannel;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.concurrent.atomic.AtomicBoolean;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies byte copying, digesting, durability evidence, and descriptor capability rejection. */
class RequestPathPublicationFileSupportTest {
  @TempDir Path root;

  @Test
  void copiesForcesAndDigestsTheExactStagedArtifact() throws Exception {
    byte[] bytes = new byte[] {1, 2, 3, 4};
    Path staged = Files.write(root.resolve("staged.xlsx"), bytes);
    Path destination = root.resolve("published.xlsx");

    try (SeekableByteChannel target =
        Files.newByteChannel(
            destination,
            StandardOpenOption.CREATE_NEW,
            StandardOpenOption.WRITE,
            StandardOpenOption.READ)) {
      RequestPathPublicationFileSupport.copyAndForce(staged, target);
    }
    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("published.xlsx", root);
        SeekableByteChannel channel = Files.newByteChannel(destination, StandardOpenOption.READ)) {
      RequestPathPublicationFileSupport.ArtifactDigest pathDigest =
          RequestPathPublicationFileSupport.digest(destination);
      RequestPathPublicationFileSupport.ArtifactDigest channelDigest =
          RequestPathPublicationFileSupport.digest(channel);
      PublicationOutcome.DurabilityEvidence durability =
          RequestPathPublicationFileSupport.forcePublishedArtifact(binding);

      assertEquals(pathDigest, channelDigest);
      assertEquals(4, pathDigest.byteSize());
      assertTrue(
          durability instanceof PublicationOutcome.DurabilityEvidence.FileAndDirectorySynced
              || durability
                  instanceof
                  PublicationOutcome.DurabilityEvidence.FileSyncedDirectoryUnestablished);
    }
    assertArrayEquals(bytes, Files.readAllBytes(destination));
  }

  @Test
  void rejectsAChannelThatCannotProvideFileForceSemanticsAndClosesIt() throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {1});
    assertThrows(
        NullPointerException.class,
        () -> RequestPathPublicationFileSupport.copyAndForce(staged, null));
    AtomicBoolean closed = new AtomicBoolean();
    assertInstanceOf(
        java.nio.file.AtomicMoveNotSupportedException.class,
        assertThrows(
            IOException.class,
            () ->
                RequestPathPublicationFileSupport.copyAndForce(
                    staged, new NonFileSeekableByteChannel(closed, false))));
    assertTrue(closed.get());
    IOException failure =
        assertThrows(
            IOException.class,
            () ->
                RequestPathPublicationFileSupport.copyAndForce(
                    staged, new NonFileSeekableByteChannel(new AtomicBoolean(), true)));
    assertEquals(1, failure.getSuppressed().length);
    assertThrows(
        IllegalArgumentException.class,
        () -> new RequestPathPublicationFileSupport.ArtifactDigest("digest", -1));
  }

  @Test
  void classifiesTheDescriptorChannelBeforePublicationBegins() throws Exception {
    Path staged = Files.write(root.resolve("descriptor-source.xlsx"), new byte[] {1});
    try (FileChannel fileChannel = FileChannel.open(staged, StandardOpenOption.READ)) {
      assertEquals(
          fileChannel,
          RequestPathPublicationFileSupport.requireFileChannel(fileChannel, staged, staged));
    }

    AtomicBoolean closed = new AtomicBoolean();
    assertThrows(
        java.nio.file.AtomicMoveNotSupportedException.class,
        () ->
            RequestPathPublicationFileSupport.requireFileChannel(
                new NonFileSeekableByteChannel(closed, false), staged, staged));
    assertTrue(closed.get());
  }

  @Test
  void fallsBackToBufferedCopyWhenTheChannelTransferReportsNoProgress() throws Exception {
    byte[] bytes = "fallback-copy".getBytes(java.nio.charset.StandardCharsets.UTF_8);
    Path staged = Files.write(root.resolve("staged-fallback.xlsx"), bytes);
    Path destination = root.resolve("fallback-destination.xlsx");

    try (FileChannel source = FileChannel.open(staged, StandardOpenOption.READ);
        FileChannel delegate =
            FileChannel.open(
                destination,
                StandardOpenOption.CREATE_NEW,
                StandardOpenOption.READ,
                StandardOpenOption.WRITE);
        ZeroProgressWritableByteChannel target = new ZeroProgressWritableByteChannel(delegate)) {
      RequestPathPublicationFileSupport.transfer(source, target);
    }

    assertArrayEquals(bytes, Files.readAllBytes(destination));
  }

  @Test
  void reportsDirectoryDurabilityAsUnestablishedWhenTheFilesystemCannotSyncIt() throws Exception {
    Path destination = Files.write(root.resolve("published.xlsx"), new byte[] {1});

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("published.xlsx", root)) {
      assertInstanceOf(
          PublicationOutcome.DurabilityEvidence.FileSyncedDirectoryUnestablished.class,
          RequestPathPublicationFileSupport.forcePublishedArtifact(
              binding,
              ignored -> {
                throw new IOException("directory fsync unavailable");
              }));
    }

    assertTrue(Files.isRegularFile(destination));
  }

  @Test
  void rejectsAStagedChannelThatEndsBeforeItsDeclaredSizeDuringFallbackCopy() throws Exception {
    Path destination = root.resolve("short-read.xlsx");

    try (FileChannel target =
        FileChannel.open(
            destination,
            StandardOpenOption.CREATE_NEW,
            StandardOpenOption.READ,
            StandardOpenOption.WRITE)) {
      IOException failure =
          assertThrows(
              IOException.class,
              () ->
                  RequestPathPublicationFileSupport.transfer(
                      new PrematureEofTransferSource(), target));

      assertEquals(
          "could not copy the complete staged workbook into publication sibling",
          failure.getMessage());
    }
  }

  @Test
  void treatsAnUnavailableMandatorySha256ProviderAsAnInternalRuntimeInvariantFailure()
      throws Exception {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                RequestPathPublicationFileSupport.digest(
                    new ByteArrayInputStream(new byte[0]),
                    algorithm -> {
                      throw new java.security.NoSuchAlgorithmException(algorithm);
                    }));

    assertEquals("SHA-256 must be available in the Java runtime", failure.getMessage());
  }
}
