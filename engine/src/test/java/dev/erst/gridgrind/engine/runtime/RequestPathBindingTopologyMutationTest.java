package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertThrows;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Exercises topology mutations at each point between descriptor binding and access. */
class RequestPathBindingTopologyMutationTest {
  @TempDir Path root;

  @Test
  @SuppressWarnings(
      "StreamResourceLeak") // The binding takes ownership of this descriptor before it can return.
  void failsClosedWhenTheExecutionRootChangesDuringBinding() throws Exception {
    Path executionRoot = Files.createDirectory(root.resolve("execution-root"));
    Path displacedRoot = root.resolve("displaced-root");

    assertThrows(
        UnsafePathAccessException.class,
        () ->
            RequestPathBinding.bindWriteTarget(
                "output.xlsx",
                executionRoot,
                directory -> {
                  Files.move(executionRoot, displacedRoot);
                  Files.createDirectory(executionRoot);
                  return Files.newDirectoryStream(displacedRoot);
                },
                SecureDirectoryStream::newByteChannel));

    Files.delete(executionRoot);
    Files.move(displacedRoot, executionRoot);
  }

  @Test
  void failsClosedWhenTheExecutionRootChangesAfterBinding() throws Exception {
    Path executionRoot = Files.createDirectory(root.resolve("execution-root"));
    Files.write(executionRoot.resolve("input.txt"), new byte[] {1});
    Path displacedRoot = root.resolve("displaced-root");

    try (RequestPathBinding binding =
        RequestPathBinding.bindExistingRead("input.txt", executionRoot)) {
      Files.move(executionRoot, displacedRoot);
      Files.createDirectory(executionRoot);

      assertThrows(UnsafePathAccessException.class, binding::openInputStream);
    }

    Files.delete(executionRoot);
    Files.move(displacedRoot, executionRoot);
  }

  @Test
  @SuppressWarnings(
      "StreamResourceLeak") // The mutating wrapper transfers the real root descriptor to binding.
  void failsClosedWhenAnAncestorChangesDuringDescriptorBinding() throws Exception {
    Path vanishedBeforeOpen = Files.createDirectory(root.resolve("vanished-before-open"));
    assertThrows(
        UnsafePathAccessException.class,
        () ->
            RequestPathBinding.bindWriteTarget(
                "vanished-before-open/output.xlsx",
                root,
                directory ->
                    new TopologyMutatingSecureDirectoryStream(
                        (SecureDirectoryStream<Path>) Files.newDirectoryStream(directory),
                        (ignored, afterChildOpen) -> {
                          if (!afterChildOpen) {
                            Files.delete(vanishedBeforeOpen);
                          }
                        }),
                SecureDirectoryStream::newByteChannel));

    Path vanishedAfterOpen = Files.createDirectory(root.resolve("vanished-after-open"));
    assertThrows(
        UnsafePathAccessException.class,
        () ->
            RequestPathBinding.bindWriteTarget(
                "vanished-after-open/output.xlsx",
                root,
                directory ->
                    new TopologyMutatingSecureDirectoryStream(
                        (SecureDirectoryStream<Path>) Files.newDirectoryStream(directory),
                        (ignored, afterChildOpen) -> {
                          if (afterChildOpen) {
                            Files.delete(vanishedAfterOpen);
                          }
                        }),
                SecureDirectoryStream::newByteChannel));

    Path replacedAncestor = Files.createDirectory(root.resolve("replaced-ancestor"));
    Path displacedAncestor = root.resolve("displaced-ancestor");
    assertThrows(
        UnsafePathAccessException.class,
        () ->
            RequestPathBinding.bindWriteTarget(
                "replaced-ancestor/output.xlsx",
                root,
                directory ->
                    new TopologyMutatingSecureDirectoryStream(
                        (SecureDirectoryStream<Path>) Files.newDirectoryStream(directory),
                        (ignored, afterChildOpen) -> {
                          if (afterChildOpen) {
                            Files.move(replacedAncestor, displacedAncestor);
                            Files.createDirectory(replacedAncestor);
                          }
                        }),
                SecureDirectoryStream::newByteChannel));
  }
}
