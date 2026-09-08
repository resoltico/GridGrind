package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import java.nio.file.StandardOpenOption;
import java.util.Set;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Exercises the descriptor binding lifecycle through real filesystem topology changes. */
class RequestPathBindingTest {
  @TempDir Path root;

  @Test
  void rejectsMissingReadLeavesAndPathsThatNameTheRootWhileCreatingContainedWriteParents()
      throws Exception {
    assertThrows(
        java.nio.file.NoSuchFileException.class,
        () -> RequestPathBinding.bindExistingRead("missing.xlsx", root));
    assertThrows(
        java.nio.file.NoSuchFileException.class,
        () -> RequestPathBinding.bindExistingRead("missing-parent/input.xlsx", root));
    assertThrows(
        UnsafePathAccessException.class, () -> RequestPathBinding.bindWriteTarget(".", root));
    try (RequestPathBinding write =
        RequestPathBinding.bindWriteTarget("missing-parent/deeper/output.xlsx", root)) {
      assertTrue(Files.isDirectory(root.resolve("missing-parent/deeper")));
      assertFalse(write.hasExistingLeaf());
    }
  }

  @Test
  void readsRegularFilesThroughBoundDescriptors() throws Exception {
    Files.write(root.resolve("input.txt"), new byte[] {4, 5});
    try (RequestPathBinding read = RequestPathBinding.bindExistingRead("input.txt", root)) {
      try (var input = read.openInputStream()) {
        assertEquals(2, input.readAllBytes().length);
      }
    }
  }

  @Test
  void rejectsLeafDeletionAndSymlinkReplacementAfterBinding() throws Exception {
    Path input = Files.write(root.resolve("input.txt"), new byte[] {4, 5});
    try (RequestPathBinding read = RequestPathBinding.bindExistingRead("input.txt", root)) {
      Files.delete(input);
      assertThrows(UnsafePathAccessException.class, read::openInputStream);
    }

    Files.write(root.resolve("input.txt"), new byte[] {4, 5});
    try (RequestPathBinding read = RequestPathBinding.bindExistingRead("input.txt", root)) {
      Files.delete(input);
      Files.createSymbolicLink(input, Path.of("other.txt"));
      assertThrows(UnsafePathAccessException.class, read::openInputStream);
    }

    Files.delete(input);
    Files.write(input, new byte[] {4, 5});
    // Keep the original inode allocated so an immediate replacement cannot reuse its file key.
    Files.createLink(root.resolve("original-input.txt"), input);
    try (RequestPathBinding read = RequestPathBinding.bindExistingRead("input.txt", root)) {
      Files.delete(input);
      Files.write(input, new byte[] {6, 7});
      assertThrows(UnsafePathAccessException.class, read::openInputStream);
    }
  }

  @Test
  void failsCreateNewWhenTheLeafAppearsAfterOutputPreflight() throws Exception {
    try (RequestPathBinding write = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      Files.write(root.resolve("output.xlsx"), new byte[] {1});
      assertThrows(UnsafePathAccessException.class, write::reverifyPublicationTarget);
    }
  }

  @Test
  @SuppressWarnings(
      "StreamResourceLeak") // The wrapper transfers delegate ownership to its close method.
  void failsClosedWhenTheFilesystemDoesNotExposeSecureDirectoryStreams() throws Exception {
    Files.write(root.resolve("input.txt"), new byte[] {1});

    assertThrows(
        UnsafePathAccessException.class,
        () ->
            RequestPathBinding.bindExistingRead(
                "input.txt",
                root,
                directory -> {
                  DirectoryStream<Path> delegate = Files.newDirectoryStream(directory);
                  return new DirectoryStream<>() {
                    @Override
                    public java.util.Iterator<Path> iterator() {
                      return delegate.iterator();
                    }

                    @Override
                    public void close() throws java.io.IOException {
                      delegate.close();
                    }
                  };
                },
                SecureDirectoryStream::newByteChannel));
  }

  @Test
  void mapsDescriptorOpenFailureToUnsafeAccessWithoutReresolvingThePath() throws Exception {
    Files.write(root.resolve("input.txt"), new byte[] {1});
    try (RequestPathBinding binding =
        RequestPathBinding.bindExistingRead(
            "input.txt",
            root,
            Files::newDirectoryStream,
            (parent, leafName, options) -> {
              throw new java.nio.file.NoSuchFileException(leafName.toString());
            })) {
      assertThrows(UnsafePathAccessException.class, binding::openInputStream);
    }
  }

  @Test
  void mapsPublishedLeafVerificationOpenFailureToUnsafeAccess() throws Exception {
    Files.write(root.resolve("published.xlsx"), new byte[] {1});

    try (RequestPathBinding binding =
        RequestPathBinding.bindWriteTarget(
            "published.xlsx",
            root,
            Files::newDirectoryStream,
            (parent, leafName, options) -> {
              throw new java.nio.file.NoSuchFileException(leafName.toString());
            })) {
      assertThrows(UnsafePathAccessException.class, binding::openPublishedReadChannel);
    }
  }

  @Test
  void mapsDetectedDescriptorLoopsToUnsafeAccessAndPreservesOrdinaryIoFailures() throws Exception {
    Files.write(root.resolve("input.txt"), new byte[] {1});
    try (RequestPathBinding binding =
        RequestPathBinding.bindExistingRead(
            "input.txt",
            root,
            Files::newDirectoryStream,
            (parent, leafName, options) -> {
              throw new java.nio.file.FileSystemLoopException(leafName.toString());
            })) {
      assertThrows(UnsafePathAccessException.class, binding::openInputStream);
    }

    try (RequestPathBinding binding =
        RequestPathBinding.bindExistingRead(
            "input.txt",
            root,
            Files::newDirectoryStream,
            (parent, leafName, options) -> {
              throw new java.nio.file.AccessDeniedException(leafName.toString());
            })) {
      assertThrows(java.nio.file.AccessDeniedException.class, binding::openInputStream);
    }

    Path target = Files.write(root.resolve("target.txt"), new byte[] {2});
    try (RequestPathBinding binding =
        RequestPathBinding.bindExistingRead(
            "input.txt",
            root,
            Files::newDirectoryStream,
            (parent, leafName, options) -> {
              Path input = root.resolve(leafName);
              Files.delete(input);
              Files.createSymbolicLink(input, target.getFileName());
              throw new java.nio.file.FileSystemLoopException(leafName.toString());
            })) {
      assertThrows(UnsafePathAccessException.class, binding::openInputStream);
    }
  }

  @Test
  @SuppressWarnings({"PMD.CloseResource", "StreamResourceLeak"})
  void propagatesAFailureClosingOneBoundDescriptor() throws Exception {
    Files.write(root.resolve("input.txt"), new byte[] {1});
    RequestPathBinding binding =
        RequestPathBinding.bindExistingRead(
            "input.txt",
            root,
            directory ->
                new CloseFailingSecureDirectoryStream(
                    (SecureDirectoryStream<Path>) Files.newDirectoryStream(directory)),
            SecureDirectoryStream::newByteChannel);

    assertThrows(java.io.IOException.class, binding::close);
  }

  @Test
  @SuppressWarnings({"PMD.CloseResource", "StreamResourceLeak"})
  void retainsEveryDescriptorCloseFailureDuringBindingAndNormalClosure() throws Exception {
    Path nested = Files.createDirectory(root.resolve("nested"));
    Files.write(nested.resolve("input.txt"), new byte[] {1});
    RequestPathBinding binding =
        RequestPathBinding.bindExistingRead(
            "nested/input.txt",
            root,
            directory ->
                new CloseFailingSecureDirectoryStream(
                    (SecureDirectoryStream<Path>) Files.newDirectoryStream(directory)),
            SecureDirectoryStream::newByteChannel);

    java.io.IOException closeFailure = assertThrows(java.io.IOException.class, binding::close);
    assertEquals(1, closeFailure.getSuppressed().length);

    java.nio.file.NoSuchFileException bindingFailure =
        assertThrows(
            java.nio.file.NoSuchFileException.class,
            () ->
                RequestPathBinding.bindExistingRead(
                    "missing.xlsx",
                    root,
                    directory ->
                        new CloseFailingSecureDirectoryStream(
                            (SecureDirectoryStream<Path>) Files.newDirectoryStream(directory)),
                    SecureDirectoryStream::newByteChannel));
    assertEquals(1, bindingFailure.getSuppressed().length);
  }

  @Test
  void managesPrivateSiblingsAndReopensThePublishedLeafThroughTheRetainedParent() throws Exception {
    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("published.xlsx", root)) {
      Path sibling = Path.of(".publication-sibling.xlsx");
      try (var channel =
          binding.openSiblingChannel(
              sibling,
              Set.of(
                  StandardOpenOption.CREATE_NEW,
                  StandardOpenOption.WRITE,
                  LinkOption.NOFOLLOW_LINKS))) {
        channel.write(java.nio.ByteBuffer.wrap(new byte[] {1, 2}));
      }
      try (var channel = binding.openSiblingReadChannel(sibling)) {
        assertEquals(2, channel.size());
      }
      assertEquals(root.resolve(sibling), binding.siblingPath(sibling));
      assertEquals(Path.of("published.xlsx"), binding.outputLeafName());
      binding.deleteSibling(sibling);

      Files.write(root.resolve("published.xlsx"), new byte[] {3, 4, 5});
      try (var input = binding.openPublishedInputStream()) {
        assertEquals(3, input.readAllBytes().length);
      }
    }
  }
}
