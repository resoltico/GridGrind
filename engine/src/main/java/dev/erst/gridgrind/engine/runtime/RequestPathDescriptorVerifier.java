package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.util.List;

/** Rechecks and releases one already-bound no-follow descriptor chain. */
final class RequestPathDescriptorVerifier {
  private RequestPathDescriptorVerifier() {}

  static void reverify(RequestPathDescriptorChain chain) throws IOException {
    try {
      reverifyDirectories(chain);
      reverifyLeaf(chain);
    } catch (NoSuchFileException exception) {
      throw topologyChanged(chain.resolvedPath(), exception);
    }
  }

  static void close(RequestPathDescriptorChain chain) throws IOException {
    IOException failure = null;
    List<RequestPathBoundDirectory> directories = chain.directories();
    for (int index = directories.size() - 1; index >= 0; index--) {
      try {
        directories.get(index).stream().close();
      } catch (IOException exception) {
        if (failure == null) {
          failure = exception;
        } else {
          failure.addSuppressed(exception);
        }
      }
    }
    if (failure != null) {
      throw failure;
    }
  }

  static RequestPathBoundDirectory parentDirectory(RequestPathDescriptorChain chain) {
    return chain.directories().getLast();
  }

  private static void reverifyDirectories(RequestPathDescriptorChain chain) throws IOException {
    List<RequestPathBoundDirectory> directories = chain.directories();
    RequestPathBoundDirectory root = directories.getFirst();
    if (!root.identity().equals(RequestPathTopology.identityOf(root.path(), true))) {
      throw topologyChanged(chain.resolvedPath());
    }
    for (int index = 1; index < directories.size(); index++) {
      reverifyDirectory(chain.resolvedPath(), directories.get(index - 1), directories.get(index));
    }
  }

  private static void reverifyDirectory(
      Path resolvedPath, RequestPathBoundDirectory parent, RequestPathBoundDirectory directory)
      throws IOException {
    if (!directory
        .identity()
        .equals(
            RequestPathTopology.identityOf(
                parent.stream(), directory.nameWithinParent(), directory.path(), true))) {
      throw topologyChanged(resolvedPath);
    }
  }

  private static void reverifyLeaf(RequestPathDescriptorChain chain) throws IOException {
    if (chain.leaf().isEmpty()) {
      reverifyAbsentLeaf(chain);
      return;
    }
    reverifyExistingLeaf(chain, chain.leaf().orElseThrow());
  }

  private static void reverifyExistingLeaf(
      RequestPathDescriptorChain chain, RequestPathTopology.ExistingFile expectedLeaf)
      throws IOException {
    if (!expectedLeaf
        .identity()
        .equals(
            RequestPathTopology.identityOf(
                parentDirectory(chain).stream(), chain.leafName(), expectedLeaf.path(), false))) {
      throw topologyChanged(chain.resolvedPath());
    }
  }

  private static void reverifyAbsentLeaf(RequestPathDescriptorChain chain) throws IOException {
    if (RequestPathTopology.optionalLeaf(
            parentDirectory(chain).stream(), chain.leafName(), chain.resolvedPath())
        .isPresent()) {
      throw new UnsafePathAccessException(
          "request-owned output path appeared after preflight: " + chain.resolvedPath());
    }
  }

  private static UnsafePathAccessException topologyChanged(Path resolvedPath) {
    return new UnsafePathAccessException(
        "request-owned path topology changed before access: " + resolvedPath);
  }

  private static UnsafePathAccessException topologyChanged(
      Path resolvedPath, NoSuchFileException cause) {
    return new UnsafePathAccessException(
        "request-owned path topology changed before access: " + resolvedPath, cause);
  }
}
