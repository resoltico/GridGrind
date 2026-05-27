package dev.erst.gridgrind.buildlogic;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.gradle.api.DefaultTask;
import org.gradle.api.file.DirectoryProperty;
import org.gradle.api.provider.ListProperty;
import org.gradle.api.tasks.Input;
import org.gradle.api.tasks.Internal;
import org.gradle.api.tasks.TaskAction;

public abstract class PruneJarOutputsTask extends DefaultTask {
  @Input
  public abstract ListProperty<String> getExpectedArchiveFileNames();

  @Internal
  public abstract DirectoryProperty getLibsDirectory();

  @TaskAction
  void pruneJarOutputs() throws IOException {
    Path libsDirectory = getLibsDirectory().get().getAsFile().toPath();
    if (!Files.isDirectory(libsDirectory)) {
      return;
    }
    Set<String> expectedArchiveNames =
        Set.copyOf(getExpectedArchiveFileNames().getOrElse(List.of()));
    try (Stream<Path> children = Files.list(libsDirectory)) {
      children
          .filter(Files::isRegularFile)
          .filter(candidate -> candidate.getFileName().toString().endsWith(".jar"))
          .filter(candidate -> !expectedArchiveNames.contains(candidate.getFileName().toString()))
          .forEach(
              candidate -> {
                try {
                  Files.deleteIfExists(candidate);
                } catch (IOException exception) {
                  throw new IllegalStateException("Failed pruning stale jar output " + candidate, exception);
                }
              });
    }
  }
}
