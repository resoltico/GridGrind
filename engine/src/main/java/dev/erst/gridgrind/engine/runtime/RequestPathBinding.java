package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.channels.Channels;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.NoSuchFileException;
import java.nio.file.OpenOption;
import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import java.nio.file.StandardOpenOption;
import java.util.Objects;
import java.util.Set;

/** Performs descriptor-relative reads and commits for one phase-four-bound request path. */
final class RequestPathBinding implements AutoCloseable {
  private final RequestPathDescriptorChain descriptorChain;
  private final ChannelOpener channels;

  private RequestPathBinding(RequestPathDescriptorChain descriptorChain, ChannelOpener channels) {
    this.descriptorChain =
        Objects.requireNonNull(descriptorChain, "descriptorChain must not be null");
    this.channels = Objects.requireNonNull(channels, "channels must not be null");
  }

  static RequestPathBinding bindExistingRead(String rawPath, Path executionRoot)
      throws IOException {
    return bind(
        rawPath,
        executionRoot,
        true,
        Files::newDirectoryStream,
        SecureDirectoryStream::newByteChannel);
  }

  static RequestPathBinding bindWriteTarget(String rawPath, Path executionRoot) throws IOException {
    return bind(
        rawPath,
        executionRoot,
        false,
        Files::newDirectoryStream,
        SecureDirectoryStream::newByteChannel);
  }

  static RequestPathBinding bindWriteTarget(
      String rawPath,
      Path executionRoot,
      DirectoryStreamFactory directoryStreams,
      ChannelOpener channels)
      throws IOException {
    return bind(rawPath, executionRoot, false, directoryStreams, channels);
  }

  static RequestPathBinding bindExistingRead(
      String rawPath,
      Path executionRoot,
      DirectoryStreamFactory directoryStreams,
      ChannelOpener channels)
      throws IOException {
    return bind(rawPath, executionRoot, true, directoryStreams, channels);
  }

  Path resolvedPath() {
    return descriptorChain.resolvedPath();
  }

  boolean hasExistingLeaf() {
    return descriptorChain.leaf().isPresent();
  }

  InputStream openInputStream() throws IOException {
    RequestPathDescriptorVerifier.reverify(descriptorChain);
    try {
      return Channels.newInputStream(
          channels.open(
              parentDirectory().stream(),
              descriptorChain.leafName(),
              Set.of(StandardOpenOption.READ, LinkOption.NOFOLLOW_LINKS)));
    } catch (IOException exception) {
      throw noFollowFailure("open request-owned file for reading", exception);
    }
  }

  void commitFrom(Path stagedFile, WorkbookArtifactWriteDisposition disposition)
      throws IOException {
    Objects.requireNonNull(stagedFile, "stagedFile must not be null");
    Objects.requireNonNull(disposition, "disposition must not be null");
    RequestPathDescriptorVerifier.reverify(descriptorChain);
    try (OutputStream output = Channels.newOutputStream(openWriteChannel(disposition))) {
      Files.copy(stagedFile, output);
    } catch (java.nio.file.FileAlreadyExistsException exception) {
      throw exception;
    } catch (IOException exception) {
      throw noFollowFailure("commit staged workbook", exception);
    }
  }

  @Override
  public void close() throws IOException {
    RequestPathDescriptorVerifier.close(descriptorChain);
  }

  private static RequestPathBinding bind(
      String rawPath,
      Path executionRoot,
      boolean requireLeaf,
      DirectoryStreamFactory directoryStreams,
      ChannelOpener channels)
      throws IOException {
    return new RequestPathBinding(
        RequestPathDescriptorBinder.bind(rawPath, executionRoot, requireLeaf, directoryStreams),
        channels);
  }

  private SeekableByteChannel openWriteChannel(WorkbookArtifactWriteDisposition disposition)
      throws IOException {
    Set<OpenOption> options =
        switch (disposition) {
          case CREATE_NEW ->
              Set.of(
                  StandardOpenOption.WRITE,
                  StandardOpenOption.CREATE_NEW,
                  LinkOption.NOFOLLOW_LINKS);
          case REPLACE_EXISTING ->
              Set.of(
                  StandardOpenOption.WRITE,
                  StandardOpenOption.CREATE,
                  StandardOpenOption.TRUNCATE_EXISTING,
                  LinkOption.NOFOLLOW_LINKS);
        };
    try {
      return channels.open(parentDirectory().stream(), descriptorChain.leafName(), options);
    } catch (java.nio.file.FileAlreadyExistsException exception) {
      throw exception;
    } catch (IOException exception) {
      throw noFollowFailure("open request-owned file for writing", exception);
    }
  }

  private RequestPathBoundDirectory parentDirectory() {
    return RequestPathDescriptorVerifier.parentDirectory(descriptorChain);
  }

  private IOException noFollowFailure(String operation, IOException cause) {
    if (cause instanceof NoSuchFileException && descriptorChain.leaf().isPresent()) {
      return new UnsafePathAccessException(
          "request-owned path disappeared before " + operation + ": " + resolvedPath(), cause);
    }
    if (cause instanceof java.nio.file.FileSystemLoopException) {
      return new UnsafePathAccessException(
          "could not safely " + operation + ": " + resolvedPath(), cause);
    }
    return cause;
  }

  /** Opens one directory stream for descriptor binding. */
  @FunctionalInterface
  interface DirectoryStreamFactory {
    /** Opens the supplied directory without imposing an application-level path policy. */
    DirectoryStream<Path> open(Path directory) throws IOException;
  }

  /** Opens one descriptor-relative file channel. */
  @FunctionalInterface
  interface ChannelOpener {
    /** Opens one leaf beneath the supplied secure parent descriptor. */
    SeekableByteChannel open(
        SecureDirectoryStream<Path> parent, Path leafName, Set<? extends OpenOption> options)
        throws IOException;
  }
}
