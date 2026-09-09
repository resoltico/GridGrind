package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.io.InputStream;
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

  /** Rechecks the bound target before a publication operation changes its leaf. */
  void reverifyPublicationTarget() throws IOException {
    RequestPathDescriptorVerifier.reverify(descriptorChain);
  }

  /** Rechecks the stable parent chain after publication intentionally replaces the target leaf. */
  void reverifyPublicationParent() throws IOException {
    RequestPathDescriptorVerifier.reverifyDirectories(descriptorChain);
  }

  /** Returns the final destination leaf name beneath the bound parent directory. */
  Path outputLeafName() {
    return descriptorChain.leafName();
  }

  /** Opens one private sibling entry beneath the bound parent without resolving a path string. */
  SeekableByteChannel openSiblingChannel(Path siblingName, Set<? extends OpenOption> options)
      throws IOException {
    Objects.requireNonNull(siblingName, "siblingName must not be null");
    Objects.requireNonNull(options, "options must not be null");
    return channels.open(parentDirectory().stream(), siblingName, options);
  }

  /** Opens one private sibling entry for no-follow verification through the retained parent. */
  SeekableByteChannel openSiblingReadChannel(Path siblingName) throws IOException {
    Objects.requireNonNull(siblingName, "siblingName must not be null");
    return channels.open(
        parentDirectory().stream(),
        siblingName,
        Set.of(StandardOpenOption.READ, LinkOption.NOFOLLOW_LINKS));
  }

  /** Opens the published leaf through the retained parent after its intentional replacement. */
  InputStream openPublishedInputStream() throws IOException {
    return Channels.newInputStream(openPublishedReadChannel());
  }

  /** Opens the published leaf through the retained parent after its intentional replacement. */
  SeekableByteChannel openPublishedReadChannel() throws IOException {
    reverifyPublicationParent();
    RequestPathTopology.identityOf(
        parentDirectory().stream(), descriptorChain.leafName(), resolvedPath(), false);
    try {
      return channels.open(
          parentDirectory().stream(),
          descriptorChain.leafName(),
          Set.of(StandardOpenOption.READ, LinkOption.NOFOLLOW_LINKS));
    } catch (IOException exception) {
      throw noFollowFailure("open published workbook for verification", exception);
    }
  }

  /** Deletes one private sibling entry beneath the retained parent descriptor. */
  void deleteSibling(Path siblingName) throws IOException {
    Objects.requireNonNull(siblingName, "siblingName must not be null");
    parentDirectory().stream().deleteFile(siblingName);
  }

  /** Resolves one private sibling only from the retained bound parent directory. */
  Path siblingPath(Path siblingName) {
    Objects.requireNonNull(siblingName, "siblingName must not be null");
    return parentDirectory().path().resolve(siblingName);
  }

  /** Returns the verified parent path used only for a qualified same-filesystem atomic move. */
  Path parentPath() {
    return parentDirectory().path();
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

  private RequestPathBoundDirectory parentDirectory() {
    return RequestPathDescriptorVerifier.parentDirectory(descriptorChain);
  }

  private IOException noFollowFailure(String operation, IOException cause) {
    if (cause instanceof NoSuchFileException) {
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
