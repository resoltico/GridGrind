package dev.erst.gridgrind.cli;

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

/**
 * Delegates no-follow directory operations while injecting one deterministic channel or close
 * fault.
 */
final class FaultInjectingSecureDirectoryStream implements SecureDirectoryStream<Path> {
  private final SecureDirectoryStream<Path> delegate;
  private final boolean endReadsEarly;
  private final boolean failClose;
  private boolean closed;

  FaultInjectingSecureDirectoryStream(
      SecureDirectoryStream<Path> delegate, boolean endReadsEarly, boolean failClose) {
    this.delegate = Objects.requireNonNull(delegate, "delegate must not be null");
    this.endReadsEarly = endReadsEarly;
    this.failClose = failClose;
  }

  @Override
  public Iterator<Path> iterator() {
    return delegate.iterator();
  }

  @Override
  public SecureDirectoryStream<Path> newDirectoryStream(Path path, LinkOption... options)
      throws IOException {
    return delegate.newDirectoryStream(path, options);
  }

  @Override
  public SeekableByteChannel newByteChannel(
      Path path,
      Set<? extends OpenOption> options,
      java.nio.file.attribute.FileAttribute<?>... attributes)
      throws IOException {
    SeekableByteChannel channel = delegate.newByteChannel(path, options, attributes);
    return endReadsEarly ? new EndOfFileChannel(channel) : channel;
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
    if (closed) {
      return;
    }
    closed = true;
    delegate.close();
    if (failClose) {
      throw new IOException("close failure");
    }
  }
}
