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

/** Secure descriptor wrapper that injects one close failure after closing its real delegate. */
final class CloseFailingSecureDirectoryStream implements SecureDirectoryStream<Path> {
  private final SecureDirectoryStream<Path> delegate;

  CloseFailingSecureDirectoryStream(SecureDirectoryStream<Path> delegate) {
    this.delegate = Objects.requireNonNull(delegate, "delegate must not be null");
  }

  @Override
  public Iterator<Path> iterator() {
    return delegate.iterator();
  }

  @Override
  public SecureDirectoryStream<Path> newDirectoryStream(Path path, LinkOption... options)
      throws IOException {
    return new CloseFailingSecureDirectoryStream(delegate.newDirectoryStream(path, options));
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
    throw new IOException("close failure");
  }
}
