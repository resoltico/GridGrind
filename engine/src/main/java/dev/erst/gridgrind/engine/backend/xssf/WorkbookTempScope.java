package dev.erst.gridgrind.engine.backend.xssf;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import java.util.function.Supplier;
import org.apache.poi.util.TempFile;
import org.apache.poi.util.TempFileCreationStrategy;

/** Routes native POI scratch through the current invocation's explicit temporary-file policy. */
public final class WorkbookTempScope {
  private WorkbookTempScope() {}

  /** Executes native work under one invocation's policy, restoring the host thread's policy. */
  public static <T> T call(WorkbookTempFileFactory factory, Supplier<T> action) {
    Objects.requireNonNull(factory, "factory must not be null");
    return TempFile.withStrategy(
        new TempFileCreationStrategy() {
          @Override
          public File createTempFile(String prefix, String suffix) throws IOException {
            return factory.createTempFile(prefix, suffix).toFile();
          }

          @Override
          public File createTempDirectory(String prefix) throws IOException {
            Path reserved = factory.createTempFile(prefix, ".directory");
            Files.delete(reserved);
            return Files.createDirectory(reserved).toFile();
          }
        },
        action);
  }
}
