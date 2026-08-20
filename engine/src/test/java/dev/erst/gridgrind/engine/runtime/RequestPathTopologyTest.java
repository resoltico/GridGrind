package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.reflect.Proxy;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.BasicFileAttributeView;
import java.nio.file.attribute.BasicFileAttributes;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Covers request-root normalization and identity facts independently from descriptor lifetime. */
class RequestPathTopologyTest {
  @TempDir Path root;

  @Test
  void resolvesContainedPathsAndRejectsEscapesAndSymlinkRoots() throws Exception {
    assertEquals(root, RequestPathTopology.rootPath(root));
    assertEquals(
        root.resolve("file.xlsx"), RequestPathTopology.resolveWithinRoot("file.xlsx", root));
    assertThrows(
        RequestPathEscapeException.class,
        () -> RequestPathTopology.resolveWithinRoot("../outside.xlsx", root));

    Path target = Files.createDirectory(root.resolve("target"));
    Path link = root.resolve("link");
    Files.createSymbolicLink(link, target.getFileName());
    assertThrows(UnsafePathAccessException.class, () -> RequestPathTopology.rootPath(link));
  }

  @Test
  void distinguishesMissingLeavesAndRejectsSymlinksAndNondirectoryParents() throws Exception {
    assertTrue(RequestPathTopology.optionalLeaf(root.resolve("missing.xlsx")).isEmpty());
    Path file = Files.write(root.resolve("file.xlsx"), new byte[] {1});
    assertFalse(RequestPathTopology.optionalLeaf(file).isEmpty());
    assertThrows(UnsafePathAccessException.class, () -> RequestPathTopology.identityOf(file, true));

    Path link = root.resolve("file-link.xlsx");
    Files.createSymbolicLink(link, file.getFileName());
    assertThrows(UnsafePathAccessException.class, () -> RequestPathTopology.optionalLeaf(link));
  }

  @Test
  void rejectsFilesystemsThatCannotProvideStableNoFollowIdentities() {
    Path path = root.resolve("synthetic.xlsx");

    assertThrows(
        UnsafePathAccessException.class,
        () ->
            RequestPathTopology.identityFromAttributes(
                path, attributes(false, false, null), false));
    assertThrows(
        UnsafePathAccessException.class,
        () -> RequestPathTopology.identityOf((BasicFileAttributeView) null, path, false));
  }

  private static BasicFileAttributes attributes(
      boolean symbolicLink, boolean directory, Object fileKey) {
    return (BasicFileAttributes)
        Proxy.newProxyInstance(
            Thread.currentThread().getContextClassLoader(),
            new Class<?>[] {BasicFileAttributes.class},
            (proxy, method, arguments) ->
                switch (method.getName()) {
                  case "isSymbolicLink" -> symbolicLink;
                  case "isDirectory" -> directory;
                  case "fileKey" -> fileKey;
                  default -> throw new UnsupportedOperationException(method.getName());
                });
  }
}
