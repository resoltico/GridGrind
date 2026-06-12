package dev.erst.gridgrind.buildlogic;

import java.io.IOException;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

public final class RepositoryJavaSourceRoots {
  private static final Set<String> EXCLUDED_DIRECTORY_NAMES =
      Set.of(".git", ".gradle", ".idea", "build", "tmp");

  private RepositoryJavaSourceRoots() {}

  public static List<Path> discover(Path repositoryRoot) throws IOException {
    Path normalizedRepositoryRoot = repositoryRoot.toAbsolutePath().normalize();
    LinkedHashSet<Path> discoveredRoots = new LinkedHashSet<>();
    Files.walkFileTree(
        normalizedRepositoryRoot,
        new SimpleFileVisitor<>() {
          @Override
          public FileVisitResult preVisitDirectory(Path directory, BasicFileAttributes attributes) {
            Path normalizedDirectory = directory.toAbsolutePath().normalize();
            if (!normalizedDirectory.equals(normalizedRepositoryRoot)
                && EXCLUDED_DIRECTORY_NAMES.contains(normalizedDirectory.getFileName().toString())) {
              return FileVisitResult.SKIP_SUBTREE;
            }
            if (isJavaSourceRoot(normalizedDirectory)) {
              discoveredRoots.add(normalizedDirectory);
              return FileVisitResult.SKIP_SUBTREE;
            }
            return FileVisitResult.CONTINUE;
          }
        });
    return discoveredRoots.stream().sorted().toList();
  }

  static List<String> relativePaths(Path repositoryRoot, List<Path> discoveredRoots) {
    Path normalizedRepositoryRoot = repositoryRoot.toAbsolutePath().normalize();
    List<String> relativePaths = new ArrayList<>(discoveredRoots.size());
    for (Path discoveredRoot : discoveredRoots) {
      relativePaths.add(
          normalizedRepositoryRoot
              .relativize(discoveredRoot.toAbsolutePath().normalize())
              .toString()
              .replace('\\', '/'));
    }
    return List.copyOf(relativePaths);
  }

  private static boolean isJavaSourceRoot(Path directory) {
    if (!"java".equals(directory.getFileName().toString())) {
      return false;
    }
    Path sourceSetDirectory = directory.getParent();
    if (sourceSetDirectory == null) {
      return false;
    }
    Path srcDirectory = sourceSetDirectory.getParent();
    return srcDirectory != null && "src".equals(srcDirectory.getFileName().toString());
  }
}
