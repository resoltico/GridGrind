package dev.erst.gridgrind.cli.io;

import java.io.IOException;
import java.lang.foreign.Arena;
import java.lang.foreign.FunctionDescriptor;
import java.lang.foreign.Linker;
import java.lang.foreign.MemoryLayout;
import java.lang.foreign.MemorySegment;
import java.lang.foreign.ValueLayout;
import java.lang.invoke.MethodHandle;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Set;

/** Qualified local no-clobber primitives, under a stable trusted parent topology. */
final class CliPublicationPrimitives {
  private static final Set<String> FILE_STORES =
      Set.of("apfs", "ext4", "xfs", "overlay", "overlayfs", "tmpfs", "fakeowner");

  private CliPublicationPrimitives() {}

  static void createFile(Path complete, Path target) throws IOException {
    requireQualified(complete, target);
    Files.createLink(target, complete);
  }

  static void createDirectory(Path complete, Path target) throws IOException {
    requireQualified(complete, target);
    if (!"UTF-8".equalsIgnoreCase(System.getProperty("sun.jnu.encoding"))) {
      throw unsupported(
          complete, target, "native recipe publication requires a UTF-8 native filename encoding");
    }
    if (!Path.of(complete.toString()).equals(complete)
        || !Path.of(target.toString()).equals(target)) {
      throw unsupported(
          complete, target, "native recipe publication requires lossless UTF-8 path encoding");
    }
    if (!CliPublicationPrimitives.class.getModule().isNativeAccessEnabled()) {
      throw unsupported(
          complete,
          target,
          "recipe directory publication requires native access; use the executable JAR or enable"
              + " native access for the CLI module");
    }
    boolean linux = "Linux".equals(System.getProperty("os.name"));
    String symbol = linux ? "renameat2" : "renamex_np";
    Linker linker = Linker.nativeLinker();
    MemorySegment address =
        linker
            .defaultLookup()
            .find(symbol)
            .orElseThrow(
                () -> unsupported(complete, target, "no atomic no-clobber directory primitive"));
    FunctionDescriptor descriptor =
        linux
            ? FunctionDescriptor.of(
                ValueLayout.JAVA_INT,
                ValueLayout.JAVA_INT,
                ValueLayout.ADDRESS,
                ValueLayout.JAVA_INT,
                ValueLayout.ADDRESS,
                ValueLayout.JAVA_INT)
            : FunctionDescriptor.of(
                ValueLayout.JAVA_INT,
                ValueLayout.ADDRESS,
                ValueLayout.ADDRESS,
                ValueLayout.JAVA_INT);
    MethodHandle rename =
        linker.downcallHandle(address, descriptor, Linker.Option.captureCallState("errno"));
    try (Arena arena = Arena.ofConfined()) {
      MemorySegment state = arena.allocate(Linker.Option.captureStateLayout());
      MemorySegment sourceName = arena.allocateFrom(complete.toString());
      MemorySegment targetName = arena.allocateFrom(target.toString());
      int result;
      try {
        // Linux AT_FDCWD=-100, RENAME_NOREPLACE=1; Darwin RENAME_EXCL=4.
        result =
            linux
                ? (int) rename.invokeExact(state, -100, sourceName, -100, targetName, 1)
                : (int) rename.invokeExact(state, sourceName, targetName, 4);
      } catch (RuntimeException | Error failure) {
        throw failure;
      } catch (Throwable failure) {
        throw new IOException("Native recipe publication invocation failed", failure);
      }
      if (result == 0) {
        return;
      }
      int error =
          (int)
              Linker.Option.captureStateLayout()
                  .varHandle(MemoryLayout.PathElement.groupElement("errno"))
                  .get(state, 0L);
      if (error == 17 || error == (linux ? 39 : 66)) {
        throw new FileAlreadyExistsException(target.toString());
      }
      throw unsupported(
          complete, target, "atomic no-clobber directory publication failed (errno=" + error + ")");
    }
  }

  private static void requireQualified(Path complete, Path target) throws IOException {
    String os = System.getProperty("os.name");
    String provider = complete.getFileSystem().provider().getClass().getName();
    boolean platform =
        "Linux".equals(os)
            ? "sun.nio.fs.LinuxFileSystemProvider".equals(provider)
            : "Mac OS X".equals(os) && "sun.nio.fs.MacOSXFileSystemProvider".equals(provider);
    if (!platform
        || !complete.getFileSystem().equals(FileSystems.getDefault())
        || !target.getFileSystem().equals(complete.getFileSystem())
        || !FILE_STORES.contains(Files.getFileStore(complete).type())) {
      throw unsupported(
          complete, target, "filesystem provider is not qualified for no-clobber publication");
    }
  }

  private static AtomicMoveNotSupportedException unsupported(
      Path source, Path target, String reason) {
    return new AtomicMoveNotSupportedException(source.toString(), target.toString(), reason);
  }
}
