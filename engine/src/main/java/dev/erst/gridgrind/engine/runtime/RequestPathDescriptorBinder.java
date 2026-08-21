package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.LinkOption;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Binds one request-owned path to a no-follow directory descriptor chain during phase four. */
final class RequestPathDescriptorBinder {
  private RequestPathDescriptorBinder() {}

  @SuppressWarnings("PMD.ExceptionAsFlowControl")
  static RequestPathDescriptorChain bind(
      String rawPath,
      Path executionRoot,
      boolean requireLeaf,
      RequestPathBinding.DirectoryStreamFactory directoryStreams) // LIM-029
      throws IOException {
    Objects.requireNonNull(rawPath, "rawPath must not be null");
    Objects.requireNonNull(executionRoot, "executionRoot must not be null");
    Path root = RequestPathTopology.rootPath(executionRoot);
    Path resolved = RequestPathTopology.resolveWithinRoot(rawPath, root);
    Path relative = root.relativize(resolved);
    if (resolved.equals(root)) {
      throw new UnsafePathAccessException(
          "request-owned path must name a file beneath the execution root");
    }

    List<RequestPathBoundDirectory> directories = new ArrayList<>();
    try {
      RequestPathBoundDirectory parent = bindRoot(root, directoryStreams, directories);
      for (int index = 0; index < relative.getNameCount() - 1; index++) {
        parent = bindAncestor(parent, relative.getName(index), requireLeaf, directories);
      }
      Optional<RequestPathTopology.ExistingFile> leaf =
          RequestPathTopology.optionalLeaf(parent.stream(), relative.getFileName(), resolved);
      if (requireLeaf && leaf.isEmpty()) {
        throw new NoSuchFileException(resolved.toString());
      }
      return new RequestPathDescriptorChain(resolved, relative.getFileName(), directories, leaf);
    } catch (IOException | RuntimeException exception) {
      closeAll(directories, exception);
      throw exception;
    }
  }

  private static RequestPathBoundDirectory bindRoot(
      Path root,
      RequestPathBinding.DirectoryStreamFactory directoryStreams,
      List<RequestPathBoundDirectory> directories)
      throws IOException {
    RequestPathTopology.Identity identity = RequestPathTopology.identityOf(root, true);
    RequestPathBoundDirectory boundRoot =
        new RequestPathBoundDirectory(
            root, Path.of("."), secureDirectory(root, directoryStreams), identity);
    directories.add(boundRoot);
    if (!identity.equals(RequestPathTopology.identityOf(root, true))) {
      throw new UnsafePathAccessException("execution root topology changed while binding: " + root);
    }
    return boundRoot;
  }

  @SuppressWarnings(
      "PMD.CloseResource") // Ownership transfers to the descriptor chain before any later failure.
  private static RequestPathBoundDirectory bindAncestor(
      RequestPathBoundDirectory parent,
      Path component,
      boolean requireLeaf,
      List<RequestPathBoundDirectory> directories)
      throws IOException {
    Path childPath = parent.path().resolve(component);
    RequestPathTopology.Identity beforeOpen =
        childIdentity(parent, component, childPath, requireLeaf);
    SecureDirectoryStream<Path> child = openChildDirectory(parent, component, childPath);
    RequestPathBoundDirectory boundChild =
        new RequestPathBoundDirectory(childPath, component, child, beforeOpen);
    directories.add(boundChild);
    RequestPathTopology.Identity afterOpen = childIdentityAfterOpen(parent, component, childPath);
    if (!beforeOpen.equals(afterOpen)) {
      throw new UnsafePathAccessException(
          "request-owned path topology changed while binding: " + childPath);
    }
    return boundChild;
  }

  private static RequestPathTopology.Identity childIdentity(
      RequestPathBoundDirectory parent, Path component, Path childPath, boolean requireLeaf)
      throws IOException {
    try {
      return RequestPathTopology.identityOf(parent.stream(), component, childPath, true);
    } catch (NoSuchFileException exception) {
      throw missingOutputParent(requireLeaf, childPath, exception);
    }
  }

  private static SecureDirectoryStream<Path> openChildDirectory(
      RequestPathBoundDirectory parent, Path component, Path childPath) throws IOException {
    try {
      return parent.stream().newDirectoryStream(component, LinkOption.NOFOLLOW_LINKS);
    } catch (NoSuchFileException exception) {
      throw topologyChangedWhileBinding(childPath, exception);
    }
  }

  private static RequestPathTopology.Identity childIdentityAfterOpen(
      RequestPathBoundDirectory parent, Path component, Path childPath) throws IOException {
    try {
      return RequestPathTopology.identityOf(parent.stream(), component, childPath, true);
    } catch (NoSuchFileException exception) {
      throw topologyChangedWhileBinding(childPath, exception);
    }
  }

  @SuppressWarnings({
    "StreamResourceLeak", // A successful secure stream transfers ownership to the descriptor chain.
    "PMD.CloseResource"
  })
  private static SecureDirectoryStream<Path> secureDirectory(
      Path directory, RequestPathBinding.DirectoryStreamFactory directoryStreams)
      throws IOException {
    DirectoryStream<Path> stream = directoryStreams.open(directory);
    if (stream instanceof SecureDirectoryStream<Path> secureDirectoryStream) {
      return secureDirectoryStream;
    }
    stream.close();
    throw new UnsafePathAccessException(
        "filesystem does not support secure no-follow directory binding: " + directory);
  }

  private static IOException missingOutputParent(
      boolean requireLeaf, Path parentPath, NoSuchFileException exception) {
    if (requireLeaf) {
      return exception;
    }
    return new IOException("request-owned output parent does not exist: " + parentPath, exception);
  }

  private static UnsafePathAccessException topologyChangedWhileBinding(
      Path path, NoSuchFileException exception) {
    return new UnsafePathAccessException(
        "request-owned path topology changed while binding: " + path, exception);
  }

  private static void closeAll(List<RequestPathBoundDirectory> directories, Exception cause) {
    for (int index = directories.size() - 1; index >= 0; index--) {
      try {
        directories.get(index).stream().close();
      } catch (IOException closeFailure) {
        cause.addSuppressed(closeFailure);
      }
    }
  }
}
