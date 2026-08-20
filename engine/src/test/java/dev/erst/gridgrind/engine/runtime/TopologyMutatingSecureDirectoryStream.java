package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.LinkOption;
import java.nio.file.OpenOption;
import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import java.nio.file.attribute.FileAttributeView;
import java.util.Iterator;
import java.util.Objects;
import java.util.Set;

/** Secure descriptor wrapper that deterministically mutates one entry at a bind-time seam. */
final class TopologyMutatingSecureDirectoryStream implements SecureDirectoryStream<Path> {
  private final SecureDirectoryStream<Path> delegate;
  private final Mutation mutation;

  TopologyMutatingSecureDirectoryStream(SecureDirectoryStream<Path> delegate, Mutation mutation) {
    this.delegate = Objects.requireNonNull(delegate, "delegate must not be null");
    this.mutation = Objects.requireNonNull(mutation, "mutation must not be null");
  }

  @Override
  public Iterator<Path> iterator() {
    return delegate.iterator();
  }

  @Override
  public SecureDirectoryStream<Path> newDirectoryStream(Path path, LinkOption... options)
      throws IOException {
    mutation.apply(path, false);
    SecureDirectoryStream<Path> child = delegate.newDirectoryStream(path, options);
    mutation.apply(path, true);
    return child;
  }

  @Override
  public SeekableByteChannel newByteChannel(
      Path path,
      Set<? extends OpenOption> options,
      java.nio.file.attribute.FileAttribute<?>... attributes)
      throws IOException {
    return delegate.newByteChannel(path, options, attributes);
  }

  @Override
  public void deleteFile(Path path) throws IOException {
    delegate.deleteFile(path);
  }

  @Override
  public void deleteDirectory(Path path) throws IOException {
    delegate.deleteDirectory(path);
  }

  @Override
  public void move(Path sourcePath, SecureDirectoryStream<Path> targetDirectory, Path targetPath)
      throws IOException {
    delegate.move(sourcePath, targetDirectory, targetPath);
  }

  @Override
  public <V extends FileAttributeView> V getFileAttributeView(Class<V> type) {
    return delegate.getFileAttributeView(type);
  }

  @Override
  public <V extends FileAttributeView> V getFileAttributeView(
      Path path, Class<V> type, LinkOption... options) {
    return delegate.getFileAttributeView(path, type, options);
  }

  @Override
  public void close() throws IOException {
    delegate.close();
  }

  /** One deterministic descriptor-entry mutation before or after a child descriptor opens. */
  @FunctionalInterface
  interface Mutation {
    /** Applies one mutation; {@code afterChildOpen} distinguishes the descriptor-open boundary. */
    void apply(Path path, boolean afterChildOpen) throws IOException;
  }
}
