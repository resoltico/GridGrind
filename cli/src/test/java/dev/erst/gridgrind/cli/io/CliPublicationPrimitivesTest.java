package dev.erst.gridgrind.cli.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Checks no-clobber publication and native-access/path-encoding rejection boundaries. */
class CliPublicationPrimitivesTest {
  @TempDir Path parent;

  @Test
  void nativePublicationRejectsAnUnqualifiedFilenameEncodingBeforeMutation() throws Exception {
    Path stage = Files.createDirectory(parent.resolve("stage"));
    Path target = parent.resolve("target");
    String previous = System.getProperty("sun.jnu.encoding");
    try {
      System.setProperty("sun.jnu.encoding", "ISO-8859-1");
      assertThrows(
          AtomicMoveNotSupportedException.class,
          () -> CliPublicationPrimitives.createDirectory(stage, target));
      assertTrue(Files.isDirectory(stage));
      assertFalse(Files.exists(target));
    } finally {
      if (previous == null) System.clearProperty("sun.jnu.encoding");
      else System.setProperty("sun.jnu.encoding", previous);
    }
  }

  @Test
  void nativePublicationRejectsAPathWhoseOsBytesCannotRoundTripThroughUtf8() throws Exception {
    org.junit.jupiter.api.Assumptions.assumeTrue(
        "Linux".equals(System.getProperty("os.name")),
        "The Linux filesystem permits non-UTF-8 names; APFS rejects them at creation");
    Process process =
        new ProcessBuilder(
                "sh", "-c", "mkdir \"$1/$(printf '\\377')\"", "raw-path-fixture", parent.toString())
            .redirectErrorStream(true)
            .start();
    java.io.Closeable terminateProcess = process::destroyForcibly;
    try (process;
        terminateProcess) {
      assertTrue(process.waitFor(30, java.util.concurrent.TimeUnit.SECONDS));
      assertEquals(
          0,
          process.exitValue(),
          new String(
              process.getInputStream().readAllBytes(), java.nio.charset.StandardCharsets.UTF_8));
    }
    Path raw;
    try (var children = Files.list(parent)) {
      raw = children.findFirst().orElseThrow();
    }
    Path stage = Files.createDirectory(parent.resolve("stage"));
    assertNotEquals(raw, Path.of(raw.toString()));
    assertThrows(
        AtomicMoveNotSupportedException.class,
        () -> CliPublicationPrimitives.createDirectory(stage, raw));
    assertTrue(Files.isDirectory(stage));
    assertTrue(Files.isDirectory(raw));
    assertFalse(Files.exists(Path.of(raw.toString())));
  }

  @Test
  void disabledNativeAccessFailsBeforePublicationWithoutVmWarnings() throws Exception {
    var builder =
        new ProcessBuilder(
                Path.of(System.getProperty("java.home"), "bin", "java").toString(),
                "-cp",
                System.getProperty("java.class.path"),
                NativeAccessDisabled.class.getName(),
                parent.toString())
            .redirectErrorStream(true);
    for (String option : new String[] {"JDK_JAVA_OPTIONS", "JAVA_TOOL_OPTIONS", "_JAVA_OPTIONS"}) {
      builder.environment().remove(option);
    }
    Process process = builder.start();
    java.io.Closeable terminateProcess = process::destroyForcibly;
    try (process;
        terminateProcess) {
      assertTrue(process.waitFor(30, java.util.concurrent.TimeUnit.SECONDS));
      assertEquals(0, process.exitValue());
      assertEquals(
          "native-disabled",
          new String(
                  process.getInputStream().readAllBytes(), java.nio.charset.StandardCharsets.UTF_8)
              .strip());
      assertFalse(Files.exists(parent.resolve("target")));
      assertTrue(Files.isDirectory(parent.resolve("stage")));
    }
  }

  /** Subprocess entrypoint deliberately launched without a native-access grant. */
  @SuppressWarnings("PMD.PublicMemberInNonPublicType")
  // The JVM subprocess needs a public probe class/main; the enclosing JUnit test is
  // package-private.
  public static final class NativeAccessDisabled {
    private NativeAccessDisabled() {}

    /** Emits the exact success witness read by the parent test after a fail-closed refusal. */
    @SuppressWarnings(
        "PMD.SystemPrintln") // Standard output is the subprocess witness protocol asserted by the
    // parent.
    public static void main(String[] arguments) throws IOException {
      Path root = Path.of(arguments[0]);
      Path stage = Files.createDirectory(root.resolve("stage"));
      try {
        CliPublicationPrimitives.createDirectory(stage, root.resolve("target"));
        throw new AssertionError("Disabled native access reached publication");
      } catch (AtomicMoveNotSupportedException expected) {
        if (!expected.getReason().contains("native access")) {
          throw expected;
        }
        System.out.println("native-disabled");
      }
    }
  }

  @Test
  void directoryCommitPublishesCompleteUnicodeWorkspaceAndNeverReplacesAnExistingLeaf()
      throws Exception {
    Path stage = Files.createDirectory(parent.resolve("stage"));
    Files.writeString(stage.resolve("request.json"), "complete request");
    Path target = parent.resolve("α workspace [#]");
    CliPublicationPrimitives.createDirectory(stage, target);
    assertFalse(Files.exists(stage));
    assertEquals("complete request", Files.readString(target.resolve("request.json")));
    for (String kind : new String[] {"empty", "nonempty", "file", "symlink"}) {
      Path source = Files.createDirectory(parent.resolve(kind + "-stage"));
      Files.writeString(source.resolve("request.json"), "own request");
      Path occupied = parent.resolve(kind);
      switch (kind) {
        case "empty" -> Files.createDirectory(occupied);
        case "nonempty" ->
            Files.writeString(Files.createDirectory(occupied).resolve("competitor"), "kept");
        case "file" -> Files.writeString(occupied, "kept");
        case "symlink" -> Files.createSymbolicLink(occupied, target);
        default -> throw new AssertionError(kind);
      }
      Object identity =
          Files.readAttributes(
                  occupied,
                  java.nio.file.attribute.BasicFileAttributes.class,
                  java.nio.file.LinkOption.NOFOLLOW_LINKS)
              .fileKey();
      assertThrows(
          FileAlreadyExistsException.class,
          () -> CliPublicationPrimitives.createDirectory(source, occupied));
      assertEquals(
          identity,
          Files.readAttributes(
                  occupied,
                  java.nio.file.attribute.BasicFileAttributes.class,
                  java.nio.file.LinkOption.NOFOLLOW_LINKS)
              .fileKey());
      assertEquals("own request", Files.readString(source.resolve("request.json")));
    }
  }

  @Test
  void linkedResponseNeverOverwritesEvenIdenticalCompetitorBytes() throws Exception {
    Path stage = Files.writeString(parent.resolve("stage"), "complete");
    Path target = Files.writeString(parent.resolve("target"), "complete");
    Object original =
        Files.readAttributes(target, java.nio.file.attribute.BasicFileAttributes.class).fileKey();
    assertThrows(
        FileAlreadyExistsException.class, () -> CliPublicationPrimitives.createFile(stage, target));
    assertEquals(
        original,
        Files.readAttributes(target, java.nio.file.attribute.BasicFileAttributes.class).fileKey());
    Files.delete(target);
    CliPublicationPrimitives.createFile(stage, target);
    Files.delete(stage);
    assertEquals("complete", Files.readString(target));
  }

  @Test
  void unqualifiedProviderFailsBeforeAnyPublication() throws Exception {
    Path archive = parent.resolve("unqualified.zip");
    try (var zip = FileSystems.newFileSystem(archive, Map.of("create", "true"))) {
      Path source = Files.writeString(zip.getPath("/stage"), "complete");
      Path target = zip.getPath("/target");
      assertThrows(
          AtomicMoveNotSupportedException.class,
          () -> CliPublicationPrimitives.createFile(source, target));
      assertFalse(Files.exists(target));
      assertTrue(Files.exists(source));
    }
  }
}
