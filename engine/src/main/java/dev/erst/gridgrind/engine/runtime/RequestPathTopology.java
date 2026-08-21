package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import java.nio.file.attribute.BasicFileAttributeView;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.Objects;
import java.util.Optional;

/** No-follow path normalization and stable filesystem-identity facts for one request root. */
final class RequestPathTopology {
  private RequestPathTopology() {}

  static Path rootPath(Path executionRoot) throws IOException {
    Path root = executionRoot.toAbsolutePath().normalize();
    rejectSymlink(root);
    return root.toRealPath(LinkOption.NOFOLLOW_LINKS);
  }

  static Path resolveWithinRoot(String rawPath, Path root) {
    Path candidate = Path.of(rawPath);
    Path resolved =
        candidate.isAbsolute()
            ? candidate.toAbsolutePath().normalize()
            : root.resolve(candidate).normalize();
    if (!resolved.startsWith(root)) {
      throw new RequestPathEscapeException("path must not escape the execution root: " + rawPath);
    }
    return resolved;
  }

  static Optional<ExistingFile> optionalLeaf(Path path) throws IOException {
    try {
      rejectSymlink(path);
      return Optional.of(new ExistingFile(path, identityOf(path, false)));
    } catch (NoSuchFileException exception) {
      return Optional.empty();
    }
  }

  static Optional<ExistingFile> optionalLeaf(
      SecureDirectoryStream<Path> parent, Path leafName, Path resolvedPath) throws IOException {
    try {
      return Optional.of(
          new ExistingFile(resolvedPath, identityOf(parent, leafName, resolvedPath, false)));
    } catch (NoSuchFileException exception) {
      return Optional.empty();
    }
  }

  static Identity identityOf(Path path, boolean requireDirectory) throws IOException {
    return identityFromAttributes(
        path,
        Files.readAttributes(path, BasicFileAttributes.class, LinkOption.NOFOLLOW_LINKS),
        requireDirectory);
  }

  static Identity identityOf(
      SecureDirectoryStream<Path> parent,
      Path entryName,
      Path resolvedPath,
      boolean requireDirectory)
      throws IOException {
    BasicFileAttributeView attributes =
        parent.getFileAttributeView(
            entryName, BasicFileAttributeView.class, LinkOption.NOFOLLOW_LINKS);
    return identityOf(attributes, resolvedPath, requireDirectory);
  }

  static Identity identityOf(
      BasicFileAttributeView attributes, Path resolvedPath, boolean requireDirectory)
      throws IOException {
    if (attributes == null) {
      throw new UnsafePathAccessException(
          "filesystem does not support no-follow descriptor attribute reads: " + resolvedPath);
    }
    return identityFromAttributes(resolvedPath, attributes.readAttributes(), requireDirectory);
  }

  static Identity identityFromAttributes(
      Path path, BasicFileAttributes attributes, boolean requireDirectory) throws IOException {
    Objects.requireNonNull(path, "path must not be null");
    Objects.requireNonNull(attributes, "attributes must not be null");
    if (attributes.isSymbolicLink()) {
      throw new UnsafePathAccessException("symbolic links are not accepted: " + path);
    }
    if (requireDirectory && !attributes.isDirectory()) {
      throw new UnsafePathAccessException("request-owned parent is not a directory: " + path);
    }
    Object fileKey = attributes.fileKey();
    if (fileKey == null) {
      throw new UnsafePathAccessException(
          "filesystem does not expose stable no-follow identities: " + path);
    }
    return new Identity(fileKey, attributes.isDirectory());
  }

  static void rejectSymlink(Path path) throws UnsafePathAccessException {
    if (Files.isSymbolicLink(path)) {
      throw new UnsafePathAccessException("symbolic links are not accepted: " + path);
    }
  }

  record ExistingFile(Path path, Identity identity) {
    ExistingFile {
      Objects.requireNonNull(path, "path must not be null");
      Objects.requireNonNull(identity, "identity must not be null");
    }
  }

  record Identity(Object fileKey, boolean directory) {
    Identity {
      Objects.requireNonNull(fileKey, "fileKey must not be null");
    }
  }
}
