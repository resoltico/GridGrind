package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence.PublicationOutcome;
import java.io.IOException;
import java.io.InputStream;
import java.nio.ByteBuffer;
import java.nio.channels.Channels;
import java.nio.channels.FileChannel;
import java.nio.channels.SeekableByteChannel;
import java.nio.channels.WritableByteChannel;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.HexFormat;
import java.util.Objects;

/** Byte-copy, identity, and force operations for a descriptor-bound publication attempt. */
final class RequestPathPublicationFileSupport {
  private RequestPathPublicationFileSupport() {}

  static void copyAndForce(Path stagedArtifact, SeekableByteChannel target) throws IOException {
    try (FileChannel targetFileChannel =
            requireFileChannel(target, stagedArtifact, stagedArtifact);
        FileChannel source = FileChannel.open(stagedArtifact, StandardOpenOption.READ)) {
      transfer(source, targetFileChannel);
      targetFileChannel.force(true);
    }
  }

  static PublicationOutcome.DurabilityEvidence forcePublishedArtifact(RequestPathBinding binding)
      throws IOException {
    return forcePublishedArtifact(binding, RequestPathPublicationFileSupport::forceDirectory);
  }

  static PublicationOutcome.DurabilityEvidence forcePublishedArtifact(
      RequestPathBinding binding, DirectoryForcer directoryForcer) throws IOException {
    Objects.requireNonNull(binding, "binding must not be null");
    Objects.requireNonNull(directoryForcer, "directoryForcer must not be null");
    try (FileChannel fileChannel =
        requireFileChannel(
            binding.openPublishedReadChannel(), binding.resolvedPath(), binding.resolvedPath())) {
      fileChannel.force(true);
    }
    try {
      directoryForcer.force(binding.parentPath());
      return new PublicationOutcome.DurabilityEvidence.FileAndDirectorySynced();
    } catch (IOException ignored) {
      return new PublicationOutcome.DurabilityEvidence.FileSyncedDirectoryUnestablished();
    }
  }

  private static void forceDirectory(Path directoryPath) throws IOException {
    try (FileChannel directory = FileChannel.open(directoryPath, StandardOpenOption.READ)) {
      directory.force(true);
    }
  }

  static ArtifactDigest digest(Path path) throws IOException {
    try (InputStream input = Files.newInputStream(path)) {
      return digest(input);
    }
  }

  static ArtifactDigest digest(SeekableByteChannel channel) throws IOException {
    try (InputStream input = Channels.newInputStream(channel)) {
      return digest(input);
    }
  }

  /**
   * Copies one staged artifact channel completely, falling back when channel transfer makes no
   * progress.
   */
  static void transfer(FileChannel source, FileChannel target) throws IOException {
    transfer(source, (WritableByteChannel) target);
  }

  static void transfer(FileChannel source, WritableByteChannel target) throws IOException {
    transfer(new FileChannelTransferSource(source), target);
  }

  static void transfer(TransferSource source, WritableByteChannel target) throws IOException {
    Objects.requireNonNull(source, "source must not be null");
    Objects.requireNonNull(target, "target must not be null");
    long sourceSize = source.size();
    long transferred = 0;
    ByteBuffer fallbackBuffer = ByteBuffer.allocate(8 * 1024);
    while (transferred < sourceSize) {
      long moved = source.transferTo(transferred, sourceSize - transferred, target);
      if (moved > 0) {
        transferred += moved;
        continue;
      }
      source.position(transferred);
      fallbackBuffer.clear();
      int read = source.read(fallbackBuffer);
      if (read < 0) {
        break;
      }
      fallbackBuffer.flip();
      while (fallbackBuffer.hasRemaining()) {
        target.write(fallbackBuffer);
      }
      transferred += read;
    }
    if (transferred != sourceSize) {
      throw new IOException("could not copy the complete staged workbook into publication sibling");
    }
  }

  /** The source operations required to transfer one staged artifact exactly. */
  abstract static class TransferSource {
    abstract long size() throws IOException;

    abstract long transferTo(long position, long count, WritableByteChannel target)
        throws IOException;

    abstract void position(long position) throws IOException;

    abstract int read(ByteBuffer destination) throws IOException;
  }

  /** Adapts the production file channel to the minimal exact-transfer source contract. */
  private static final class FileChannelTransferSource extends TransferSource {
    private final FileChannel channel;

    private FileChannelTransferSource(FileChannel channel) {
      this.channel = Objects.requireNonNull(channel);
    }

    @Override
    long size() throws IOException {
      return channel.size();
    }

    @Override
    long transferTo(long position, long count, WritableByteChannel target) throws IOException {
      return channel.transferTo(position, count, target);
    }

    @Override
    void position(long position) throws IOException {
      channel.position(position);
    }

    @Override
    int read(ByteBuffer destination) throws IOException {
      return channel.read(destination);
    }
  }

  private static ArtifactDigest digest(InputStream input) throws IOException {
    return digest(input, MessageDigest::getInstance);
  }

  static ArtifactDigest digest(InputStream input, DigestFactory digestFactory) throws IOException {
    MessageDigest digest = sha256Digest(digestFactory);
    long byteSize = 0;
    byte[] buffer = new byte[8 * 1024];
    while (true) {
      int read = input.read(buffer);
      if (read < 0) {
        break;
      }
      digest.update(buffer, 0, read);
      byteSize += read;
    }
    return new ArtifactDigest(HexFormat.of().formatHex(digest.digest()), byteSize);
  }

  @SuppressWarnings(
      "PatternMatchingInstanceof") // Keeps JaCoCo attribution for the throwing branch.
  static FileChannel requireFileChannel(SeekableByteChannel channel, Path source, Path destination)
      throws IOException {
    if (!(channel instanceof FileChannel)) {
      try (channel) {
        Objects.requireNonNull(channel, "channel must not be null");
        throw new AtomicMoveNotSupportedException(
            source.toString(),
            destination.toString(),
            "filesystem does not expose a forceable descriptor-relative publication channel");
      }
    }
    return (FileChannel) channel;
  }

  private static MessageDigest sha256Digest(DigestFactory digestFactory) {
    Objects.requireNonNull(digestFactory, "digestFactory must not be null");
    try {
      return digestFactory.create("SHA-256");
    } catch (NoSuchAlgorithmException exception) {
      throw new IllegalStateException("SHA-256 must be available in the Java runtime", exception);
    }
  }

  record ArtifactDigest(String sha256, long byteSize) {
    ArtifactDigest {
      Objects.requireNonNull(sha256, "sha256 must not be null");
      if (byteSize < 0) {
        throw new IllegalArgumentException("byteSize must be >= 0");
      }
    }
  }

  /**
   * Synchronizes one publication parent directory after the artifact data reaches stable storage.
   */
  @FunctionalInterface
  interface DirectoryForcer {
    /**
     * Forces the supplied parent directory entry to durable storage.
     *
     * @param directoryPath verified destination parent to synchronize
     * @throws IOException when the filesystem cannot establish directory durability
     */
    void force(Path directoryPath) throws IOException;
  }

  /** Creates one message digest by algorithm name for an artifact identity calculation. */
  @FunctionalInterface
  interface DigestFactory {
    /**
     * Creates the requested digest implementation.
     *
     * @param algorithm JCA digest algorithm name
     * @return initialized message digest
     * @throws NoSuchAlgorithmException when the runtime lacks the requested algorithm
     */
    MessageDigest create(String algorithm) throws NoSuchAlgorithmException;
  }
}
