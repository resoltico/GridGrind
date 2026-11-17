package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.engine.backend.streaming.StreamingWorkbookWriter;
import java.io.File;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import org.apache.poi.util.DefaultTempFileCreationStrategy;
import org.apache.poi.util.TempFile;
import org.apache.poi.util.TempFileCreationStrategy;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.api.parallel.Isolated;

/** Native scoped policy must neither replace host policy nor escape the admitted root. */
@Isolated("One case deliberately installs a host-global POI policy")
class WorkbookTempScopeTest {
  @TempDir Path root;

  @Test
  void preservesHostGlobalAndThreadPoliciesOnSuccessAndFailure() throws Exception {
    Path global = Files.createDirectory(root.resolve("global"));
    Path local = Files.createDirectory(root.resolve("local"));
    Path run = Files.createDirectory(root.resolve("run"));
    TempFile.setTempFileCreationStrategy(strategy(global));
    try {
      Class.forName(WorkbookTempScope.class.getName());
      assertEquals(global, nativeFile().getParent());
      assertEquals(
          run,
          WorkbookTempScope.call(WorkbookTempFileFactory.rooted(run), this::nativeFile)
              .getParent());
      assertEquals(global, nativeFile().getParent());
      TempFile.withStrategy(
          strategy(local),
          () -> {
            assertEquals(
                run,
                WorkbookTempScope.call(WorkbookTempFileFactory.rooted(run), this::nativeFile)
                    .getParent());
            assertEquals(local, nativeFile().getParent());
            assertThrows(
                IllegalStateException.class,
                () ->
                    WorkbookTempScope.call(
                        WorkbookTempFileFactory.rooted(run),
                        () -> {
                          assertEquals(run, nativeFile().getParent());
                          throw new IllegalStateException("synthetic native failure");
                        }));
            assertEquals(local, nativeFile().getParent());
            return true;
          });
      assertEquals(global, nativeFile().getParent());
    } finally {
      TempFile.setTempFileCreationStrategy(new DefaultTempFileCreationStrategy());
    }
  }

  @Test
  void nestedScopesRestoreTheOuterRunThenHostPolicy() throws Exception {
    Path host = Files.createDirectory(root.resolve("host"));
    Path outer = Files.createDirectory(root.resolve("outer"));
    Path inner = Files.createDirectory(root.resolve("inner"));
    TempFile.withStrategy(
        strategy(host),
        () -> {
          WorkbookTempScope.call(
              WorkbookTempFileFactory.rooted(outer),
              () -> {
                assertEquals(outer, nativeFile().getParent());
                assertThrows(
                    IllegalArgumentException.class,
                    () ->
                        WorkbookTempScope.call(
                            WorkbookTempFileFactory.rooted(inner),
                            () -> {
                              assertEquals(inner, nativeDirectory().getParent());
                              throw new IllegalArgumentException("inner failure");
                            }));
                assertEquals(outer, nativeFile().getParent());
                return true;
              });
          assertEquals(host, nativeFile().getParent());
          return true;
        });
  }

  @Test
  void concurrentRunsOverrideAndRestoreOnlyTheirOwnHostThreadPolicy() throws Exception {
    CountDownLatch entered = new CountDownLatch(2);
    try (var executor = Executors.newFixedThreadPool(2)) {
      var first = executor.submit(() -> concurrentRun("first", entered));
      var second = executor.submit(() -> concurrentRun("second", entered));
      assertTrue(first.get(30, TimeUnit.SECONDS));
      assertTrue(second.get(30, TimeUnit.SECONDS));
    }
  }

  private boolean concurrentRun(String name, CountDownLatch entered) throws IOException {
    Path host = Files.createDirectory(root.resolve(name + "-host"));
    Path run = Files.createDirectory(root.resolve(name + "-run"));
    return TempFile.withStrategy(
        strategy(host),
        () -> {
          WorkbookTempScope.call(
              WorkbookTempFileFactory.rooted(run),
              () -> {
                entered.countDown();
                try {
                  assertTrue(entered.await(10, TimeUnit.SECONDS));
                } catch (InterruptedException exception) {
                  Thread.currentThread().interrupt();
                  throw new IllegalStateException(exception);
                }
                assertEquals(run, nativeFile().getParent());
                return true;
              });
          assertEquals(host, nativeFile().getParent());
          return true;
        });
  }

  @Test
  void realSxssfScratchUsesRunRootEvenWhenHostThreadAlreadyHasAPolicy() throws Exception {
    Path host = Files.createDirectory(root.resolve("host"));
    Path run = Files.createDirectory(root.resolve("run"));
    TempFile.withStrategy(
        strategy(host),
        () ->
            WorkbookTempScope.call(
                WorkbookTempFileFactory.rooted(run),
                () -> {
                  try (var writer = new StreamingWorkbookWriter()) {
                    writer.apply(
                        new ResolvedStep(
                            new SheetSelector.ByName(new SheetName("Data")),
                            new WorkbookMutationAction.EnsureSheet()));
                    try (var paths = Files.walk(run)) {
                      assertTrue(
                          paths.anyMatch(
                              path -> path.getFileName().toString().startsWith("poi-sxssf-sheet")));
                    }
                    try (var paths = Files.list(host)) {
                      assertEquals(List.of(), paths.toList());
                    }
                  } catch (IOException exception) {
                    throw new UncheckedIOException(exception);
                  }
                  return true;
                }));
    try (var paths = Files.list(run)) {
      assertFalse(
          paths.anyMatch(path -> path.getFileName().toString().startsWith("poi-sxssf-sheet")));
    }
  }

  private Path nativeFile() {
    try {
      return TempFile.createTempFile("native-", ".tmp").toPath();
    } catch (IOException exception) {
      throw new UncheckedIOException(exception);
    }
  }

  private Path nativeDirectory() {
    try {
      return TempFile.createTempDirectory("native-").toPath();
    } catch (IOException exception) {
      throw new UncheckedIOException(exception);
    }
  }

  private static TempFileCreationStrategy strategy(Path directory) {
    return new TempFileCreationStrategy() {
      @Override
      public File createTempFile(String prefix, String suffix) throws IOException {
        return Files.createTempFile(directory, prefix, suffix).toFile();
      }

      @Override
      public File createTempDirectory(String prefix) throws IOException {
        return Files.createTempDirectory(directory, prefix).toFile();
      }
    };
  }
}
