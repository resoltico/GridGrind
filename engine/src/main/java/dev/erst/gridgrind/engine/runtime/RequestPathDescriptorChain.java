package dev.erst.gridgrind.engine.runtime;

import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Immutable descriptor-chain facts retained from phase-four request-path binding. */
record RequestPathDescriptorChain(
    Path resolvedPath,
    Path leafName,
    List<RequestPathBoundDirectory> directories,
    Optional<RequestPathTopology.ExistingFile> leaf) {
  RequestPathDescriptorChain {
    Objects.requireNonNull(resolvedPath, "resolvedPath must not be null");
    Objects.requireNonNull(leafName, "leafName must not be null");
    directories = List.copyOf(Objects.requireNonNull(directories, "directories must not be null"));
    Objects.requireNonNull(leaf, "leaf must not be null");
  }
}

/** One directory descriptor plus its no-follow identity and entry name within the parent. */
record RequestPathBoundDirectory(
    Path path,
    Path nameWithinParent,
    SecureDirectoryStream<Path> stream,
    RequestPathTopology.Identity identity) {
  RequestPathBoundDirectory {
    Objects.requireNonNull(path, "path must not be null");
    Objects.requireNonNull(nameWithinParent, "nameWithinParent must not be null");
    Objects.requireNonNull(stream, "stream must not be null");
    Objects.requireNonNull(identity, "identity must not be null");
  }
}
