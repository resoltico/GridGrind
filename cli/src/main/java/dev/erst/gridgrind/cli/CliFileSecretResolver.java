package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.SecretReference;
import dev.erst.gridgrind.engine.api.GridGrindSecretResolver;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.channels.SeekableByteChannel;
import java.nio.charset.CharacterCodingException;
import java.nio.charset.CodingErrorAction;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.OpenOption;
import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import java.nio.file.attribute.BasicFileAttributeView;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.Objects;
import java.util.Set;

/** Resolves one approved secret file through a no-follow directory capability. */
final class CliFileSecretResolver implements GridGrindSecretResolver, AutoCloseable {
  private static final int MAX_SECRET_BYTES = 65_536;
  private final SecureDirectoryStream<Path> directory;

  static CliFileSecretResolver open(Path root) throws IOException {
    Path normalizedRoot =
        Objects.requireNonNull(root, "root must not be null").toAbsolutePath().normalize();
    if (Files.isSymbolicLink(normalizedRoot)) {
      throw new IOException("secret-provider root must not be a symbolic link");
    }
    return new CliFileSecretResolver(normalizedRoot);
  }

  private CliFileSecretResolver(Path normalizedRoot) throws IOException {
    this(openProviderDirectory(normalizedRoot));
  }

  private static DirectoryStream<Path> openProviderDirectory(Path normalizedRoot)
      throws IOException {
    Path root = Objects.requireNonNull(normalizedRoot, "normalizedRoot must not be null");
    return root.getFileSystem().provider().newDirectoryStream(root, entry -> true);
  }

  /**
   * Owns one already-open provider descriptor and rejects any filesystem without no-follow access.
   */
  CliFileSecretResolver(DirectoryStream<Path> directory) throws IOException {
    this.directory =
        requireSecureDirectory(Objects.requireNonNull(directory, "directory must not be null"));
  }

  private static SecureDirectoryStream<Path> requireSecureDirectory(DirectoryStream<Path> candidate)
      throws IOException {
    if (!(candidate instanceof SecureDirectoryStream<Path> secureDirectory)) {
      candidate.close();
      throw new IOException(
          "secret-provider filesystem does not support no-follow directory access");
    }
    return secureDirectory;
  }

  @Override
  public char[] resolve(SecretReference reference) {
    Objects.requireNonNull(reference, "reference must not be null");
    try {
      Path leaf = Path.of(reference.id());
      verifyRegularFile(leaf);
      return decode(readProviderLeaf(leaf));
    } catch (IOException exception) {
      throw new CliSecretReadException(reference.id(), exception);
    }
  }

  private void verifyRegularFile(Path leaf) throws IOException {
    BasicFileAttributes attributes =
        directory
            .getFileAttributeView(leaf, BasicFileAttributeView.class, LinkOption.NOFOLLOW_LINKS)
            .readAttributes();
    if (!attributes.isRegularFile()) {
      throw new IllegalArgumentException("secret reference does not name a regular provider file");
    }
  }

  private ByteBuffer readProviderLeaf(Path leaf) throws IOException {
    try (SeekableByteChannel channel =
        directory.newByteChannel(
            leaf,
            Set.<OpenOption>of(java.nio.file.StandardOpenOption.READ, LinkOption.NOFOLLOW_LINKS))) {
      if (channel.size() > MAX_SECRET_BYTES) {
        throw new IllegalArgumentException("secret provider file exceeds 65536 bytes");
      }
      ByteBuffer bytes = ByteBuffer.allocate(Math.toIntExact(channel.size()));
      while (bytes.hasRemaining() && channel.read(bytes) >= 0) {
        // Read the sole explicitly named provider leaf to completion.
      }
      if (bytes.hasRemaining()) {
        throw new IllegalArgumentException("secret provider file ended before its declared size");
      }
      bytes.flip();
      return bytes;
    }
  }

  private static char[] decode(ByteBuffer bytes) {
    try {
      return StandardCharsets.UTF_8
          .newDecoder()
          .onMalformedInput(CodingErrorAction.REPORT)
          .onUnmappableCharacter(CodingErrorAction.REPORT)
          .decode(bytes)
          .toString()
          .toCharArray();
    } catch (CharacterCodingException exception) {
      throw new IllegalArgumentException("secret provider file is not valid UTF-8", exception);
    }
  }

  @Override
  public void close() throws IOException {
    directory.close();
  }
}
