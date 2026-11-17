package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import java.nio.file.attribute.BasicFileAttributeView;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import org.jspecify.annotations.Nullable;

/** Verified directory-entry spelling, shared by containment, authority and snapshot identity. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_REQUEST_OWNED_PATH_CONTAINMENT
})
final class RequestPathIdentity {
  private final Path declaredRoot;
  private @Nullable Root admittedRoot;
  private List<EntryPin> qualifiedRootPrefix = List.of();

  // One invocation owns this metadata proof; no inode is an authority/cache key.
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private final Map<Path, Path> completeSpellings = new LinkedHashMap<>();

  private final Set<Path> rootSpellings = new LinkedHashSet<>();

  RequestPathIdentity(Path root) {
    declaredRoot = Objects.requireNonNull(root).toAbsolutePath().normalize();
  }

  Path rootPath() throws IOException {
    return root().path();
  }

  RequestPathTopology.Identity rootIdentity() throws IOException {
    return root().identity();
  }

  void verifyRoot() throws IOException {
    Root root = root();
    for (EntryPin pin : qualifiedRootPrefix) {
      if (!pin.stamp().equals(stamp(pin.path()))) {
        throw new UnsafePathAccessException("execution root topology changed: " + declaredRoot);
      }
    }
    if (!root.identity().equals(RequestPathTopology.identityOf(root.path(), true))) {
      throw new UnsafePathAccessException("execution root topology changed: " + declaredRoot);
    }
  }

  @SuppressWarnings(
      "PMD.CloseResource") // Every opened metadata descriptor transfers to DirectoryChain.
  Path resolve(String rawPath) throws IOException {
    Root root = root();
    Path supplied = declaredRoot.getFileSystem().getPath(rawPath);
    Path query = (supplied.isAbsolute() ? supplied : declaredRoot.resolve(supplied)).normalize();
    if (!supplied.isAbsolute() && !query.startsWith(declaredRoot)) throw escape(rawPath);
    verifyRoot();
    Path cached = completeSpellings.get(query);
    if (cached != null) return cached;
    Path relative = null;
    for (Path spelling : rootSpellings) {
      if (query.startsWith(spelling)) {
        relative = spelling.relativize(query);
        break;
      }
    }
    if (relative == null) {
      if (!Objects.equals(query.getRoot(), root.path().getRoot())
          || query.getNameCount() < root.path().getNameCount()) throw escape(rawPath);
      Path prefix = query.getRoot().resolve(query.subpath(0, root.path().getNameCount()));
      try {
        qualifyRootSpelling(prefix, root);
      } catch (NoSuchFileException absentPrefix) {
        throw new RequestPathEscapeException(
            "path must not escape the execution root: " + rawPath, absentPrefix);
      }
      rootSpellings.add(prefix);
      relative = prefix.relativize(query);
    }
    verifyRoot();
    try (DirectoryChain streams = new DirectoryChain()) {
      SecureDirectoryStream<Path> parent = open(root.path());
      streams.add(parent);
      verifyOpenedRoot(parent, root);
      Path resolved = root.path();
      if (relative.toString().isEmpty()) return resolved;
      for (int index = 0; index < relative.getNameCount(); index++) {
        Path component = relative.getName(index);
        Entry entry;
        try {
          entry = entry(parent, component, false);
        } catch (NoSuchFileException absent) {
          Path absentPath = resolved.resolve(relative.subpath(index, relative.getNameCount()));
          completeSpellings.put(query, absentPath);
          return absentPath;
        }
        resolved = resolved.resolve(entry.name());
        if (index == relative.getNameCount() - 1) {
          completeSpellings.put(query, resolved);
          return resolved;
        }
        if (!entry.stamp().directory())
          throw new UnsafePathAccessException("request parent is not a directory: " + resolved);
        SecureDirectoryStream<Path> child =
            parent.newDirectoryStream(entry.name(), LinkOption.NOFOLLOW_LINKS);
        streams.add(child);
        if (!entry.stamp().equals(stamp(parent, entry.name()))
            || !entry.stamp().equals(selfStamp(child)))
          throw new UnsafePathAccessException("request path topology changed: " + resolved);
        parent = child;
      }
      throw new IllegalStateException("A nonempty request path must resolve an entry");
    }
  }

  /** Unused declarations cannot authorize by inode, create entries, or poison another match. */
  boolean matchesGrant(Path grant, Path requested) {
    if (grant.equals(requested)) return true;
    try {
      return resolve(grant.toString()).equals(requested);
    } catch (IOException | RequestPathEscapeException unavailableDeclaration) {
      return false;
    }
  }

  private Root root() throws IOException {
    if (admittedRoot != null) return admittedRoot;
    RequestPathTopology.Identity before = RequestPathTopology.identityOf(declaredRoot, true);
    Root result = new Root(declaredRoot, before);
    try (SecureDirectoryStream<Path> stream = open(declaredRoot)) {
      verifyOpenedRoot(stream, result);
      if (!before.equals(RequestPathTopology.identityOf(declaredRoot, true))) {
        throw new UnsafePathAccessException(
            "execution root changed during admission: " + declaredRoot);
      }
    }
    admittedRoot = result;
    rootSpellings.add(declaredRoot);
    return result;
  }

  private void qualifyRootSpelling(Path supplied, Root root) throws IOException {
    // Only an unknown absolute alias needs directory-entry evidence above the host's root.
    // The fixed host root remains the resource-key anchor, including on execute-only ancestors.
    List<EntryPin> expected = qualifyPrefix(declaredRoot);
    List<EntryPin> actual = qualifyPrefix(supplied);
    if (expected.size() != actual.size()) throw escape(supplied.toString());
    for (int index = 0; index < expected.size(); index++) {
      EntryPin left = expected.get(index);
      EntryPin right = actual.get(index);
      if (!left.path().getFileName().equals(right.path().getFileName())
          || !left.stamp().equals(right.stamp())) throw escape(supplied.toString());
    }
    qualifiedRootPrefix = expected;
    verifyRoot();
    if (!root.identity().equals(RequestPathTopology.identityOf(supplied, true))) {
      throw new UnsafePathAccessException("execution root alias changed during proof: " + supplied);
    }
  }

  private static List<EntryPin> qualifyPrefix(Path supplied) throws IOException {
    Path stored = supplied.getRoot();
    List<EntryPin> pins = new ArrayList<>();
    for (int index = 0; index < supplied.getNameCount(); index++) {
      try (SecureDirectoryStream<Path> parent = open(stored)) {
        // Only explicitly host-supplied above-root topology may contain symlinks (e.g. /var).
        Entry value = entry(parent, supplied.getName(index), index < supplied.getNameCount() - 1);
        stored = stored.resolve(value.name());
        pins.add(new EntryPin(stored, value.stamp()));
      }
    }
    return List.copyOf(pins);
  }

  /** Looks up an admitted spelling without touching the original input again. */
  Path previouslyResolved(String rawPath) {
    Path supplied = declaredRoot.getFileSystem().getPath(rawPath);
    Path query = (supplied.isAbsolute() ? supplied : declaredRoot.resolve(supplied)).normalize();
    Path resolved = completeSpellings.get(query);
    if (resolved == null)
      throw new IllegalStateException("request path was not resolved during admission: " + rawPath);
    return resolved;
  }

  private static void verifyOpenedRoot(SecureDirectoryStream<Path> stream, Root root)
      throws IOException {
    Stamp opened = selfStamp(stream);
    if (!opened.key().equals(root.identity().fileKey()) || !opened.directory()) {
      throw new UnsafePathAccessException("execution root changed while opening: " + root.path());
    }
  }

  private static Entry entry(
      SecureDirectoryStream<Path> parent, Path requested, boolean allowSymbolic)
      throws IOException {
    Stamp before = stamp(parent, requested);
    if (before.symbolic() && !allowSymbolic)
      throw new UnsafePathAccessException("symbolic links are not accepted: " + requested);
    Path unique = null;
    int candidates = 0;
    try (SecureDirectoryStream<Path> listing =
        parent.newDirectoryStream(
            requested.getFileSystem().getPath("."), LinkOption.NOFOLLOW_LINKS)) {
      for (Path path : listing) {
        Path name = path.getFileName();
        if (name.equals(requested)) {
          if (!before.equals(stamp(parent, name)))
            throw new UnsafePathAccessException(
                "entry changed during spelling proof: " + requested);
          return new Entry(name, before);
        }
        Stamp candidate;
        try {
          candidate = stamp(parent, name);
        } catch (NoSuchFileException unrelatedEntryRemoved) {
          continue;
        }
        if (before.equals(candidate)) {
          unique = name;
          candidates++;
        }
      }
    }
    if (candidates != 1 || unique == null)
      throw new UnsafePathAccessException(
          "filesystem cannot prove an unambiguous entry spelling; use its exact stored name: "
              + requested);
    if (!before.equals(stamp(parent, requested)) || !before.equals(stamp(parent, unique))) {
      throw new UnsafePathAccessException("entry changed during spelling proof: " + requested);
    }
    return new Entry(unique, before);
  }

  private static Stamp stamp(Path path) throws IOException {
    return stamp(Files.readAttributes(path, BasicFileAttributes.class, LinkOption.NOFOLLOW_LINKS));
  }

  private static Stamp stamp(SecureDirectoryStream<Path> parent, Path entry) throws IOException {
    BasicFileAttributeView view =
        parent.getFileAttributeView(entry, BasicFileAttributeView.class, LinkOption.NOFOLLOW_LINKS);
    if (view == null)
      throw new UnsafePathAccessException("filesystem lacks no-follow entry attributes");
    return stamp(view.readAttributes());
  }

  private static Stamp selfStamp(SecureDirectoryStream<Path> stream) throws IOException {
    BasicFileAttributeView view = stream.getFileAttributeView(BasicFileAttributeView.class);
    if (view == null) throw new UnsafePathAccessException("filesystem lacks descriptor attributes");
    return stamp(view.readAttributes());
  }

  private static Stamp stamp(BasicFileAttributes attributes) throws IOException {
    Object key = attributes.fileKey();
    if (key == null)
      throw new UnsafePathAccessException("filesystem lacks stable entry identities");
    return new Stamp(key, attributes.isDirectory(), attributes.isSymbolicLink());
  }

  @SuppressWarnings(
      "PMD.CloseResource") // A successful secure stream transfers to its explicit caller.
  private static SecureDirectoryStream<Path> open(Path path) throws IOException {
    DirectoryStream<Path> stream = Files.newDirectoryStream(path);
    if (stream instanceof SecureDirectoryStream<Path> secure) return secure;
    stream.close();
    throw new UnsafePathAccessException("filesystem lacks secure directory access: " + path);
  }

  /** Temporary metadata descriptors close with standard primary/suppressed failure semantics. */
  private static final class DirectoryChain implements AutoCloseable {
    private final List<SecureDirectoryStream<Path>> streams = new ArrayList<>();

    void add(SecureDirectoryStream<Path> stream) {
      streams.add(stream);
    }

    @Override
    public void close() throws IOException {
      IOException failure = null;
      for (int index = streams.size() - 1; index >= 0; index--)
        try {
          streams.get(index).close();
        } catch (IOException problem) {
          if (failure == null) failure = problem;
          else if (!failure.equals(problem)) failure.addSuppressed(problem);
        }
      if (failure != null) throw failure;
    }
  }

  private static RequestPathEscapeException escape(String path) {
    return new RequestPathEscapeException("path must not escape the execution root: " + path);
  }

  private record Stamp(Object key, boolean directory, boolean symbolic) {}

  private record Entry(Path name, Stamp stamp) {}

  private record EntryPin(Path path, Stamp stamp) {}

  private record Root(Path path, RequestPathTopology.Identity identity) {}
}
