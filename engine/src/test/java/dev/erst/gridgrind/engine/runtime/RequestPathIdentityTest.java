package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Native directory-entry evidence, independent of presentation spelling and file contents. */
class RequestPathIdentityTest {
  @TempDir Path root;

  @Test
  void exactHardlinksRemainDifferentResourcesAndUnusedGrantsStayInert() throws Exception {
    Path first = Files.writeString(root.resolve("First.txt"), "one");
    Path second = Files.createLink(root.resolve("Second.txt"), first);
    RequestPathIdentity identity = new RequestPathIdentity(root);
    assertEquals(first, identity.resolve("First.txt"));
    assertEquals(second, identity.resolve("Second.txt"));
    assertFalse(identity.matchesGrant(first, second));
    assertFalse(identity.matchesGrant(root.resolve("absent.txt"), second));
    assertFalse(identity.matchesGrant(root.getParent().resolve("outside.txt"), second));
    assertTrue(identity.matchesGrant(second, second));
  }

  @Test
  void nativeCaseAndUnicodeAliasesUseStoredEntryNamesOnlyWhenTheHostResolvesThem()
      throws Exception {
    Path file = Files.writeString(root.resolve("Caf\u00e9.txt"), "one");
    RequestPathIdentity identity = new RequestPathIdentity(root);
    Path alias = root.resolve("CAFE\u0301.TXT");
    if (Files.exists(alias)) {
      assertEquals(file, identity.resolve(alias.toString()));
      assertTrue(identity.matchesGrant(alias, file));
      Files.createLink(root.resolve("Other.txt"), file);
      RequestPathIdentity fresh = new RequestPathIdentity(root);
      assertThrows(UnsafePathAccessException.class, () -> fresh.resolve(alias.toString()));
      assertEquals(file, fresh.resolve(file.toString()));
    } else {
      assertNotEquals(file, identity.resolve(alias.toString()));
      assertFalse(identity.matchesGrant(alias, file));
    }
  }

  @Test
  void rootAliasesRequireTheSameOrderedEntryChain() throws Exception {
    Path directory = Files.createDirectory(root.resolve("MixedRoot"));
    Path source = Files.writeString(directory.resolve("Input.txt"), "one");
    RequestPathIdentity identity = new RequestPathIdentity(directory);
    Path alias = root.resolve("MIXEDROOT").resolve("INPUT.TXT");
    if (Files.exists(alias)) assertEquals(source, identity.resolve(alias.toString()));
    else assertThrows(RequestPathEscapeException.class, () -> identity.resolve(alias.toString()));
    Path symbolic = Files.createSymbolicLink(root.resolve("alternate"), directory.getFileName());
    assertThrows(
        UnsafePathAccessException.class,
        () -> identity.resolve(symbolic.resolve("Input.txt").toString()));
  }

  @Test
  void absentSuffixesKeepExactNamesAndDescendantSymlinksRefuse() throws Exception {
    RequestPathIdentity identity = new RequestPathIdentity(root);
    assertEquals(root.resolve("New/Book.xlsx"), identity.resolve("New/Book.xlsx"));
    assertFalse(
        identity.matchesGrant(root.resolve("New/book.xlsx"), root.resolve("New/Book.xlsx")));
    Path target = Files.writeString(root.resolve("target"), "one");
    Files.createSymbolicLink(root.resolve("link"), target.getFileName());
    assertThrows(UnsafePathAccessException.class, () -> identity.resolve("link"));
    assertThrows(RequestPathEscapeException.class, () -> identity.resolve("../outside"));
  }

  @Test
  void canonicalAccessNeedsOnlyTraversalPermissionAboveTheAdmittedRoot() throws Exception {
    Path ancestor = Files.createDirectory(root.resolve("traverse-only"));
    Path work = Files.createDirectory(ancestor.resolve("Work"));
    Path source = Files.writeString(work.resolve("input"), "one");
    var permissions = Files.getPosixFilePermissions(ancestor);
    try {
      Files.setPosixFilePermissions(
          ancestor, java.util.Set.of(java.nio.file.attribute.PosixFilePermission.OWNER_EXECUTE));
      var identity = new RequestPathIdentity(work);
      assertEquals(source, identity.resolve("input"));
      assertEquals(source, identity.resolve(source.toString()));
      Path alias = ancestor.resolve("WORK/input");
      if (!Files.isReadable(ancestor) && Files.exists(alias)) {
        assertThrows(
            java.nio.file.AccessDeniedException.class, () -> identity.resolve(alias.toString()));
      }
    } finally {
      Files.setPosixFilePermissions(ancestor, permissions);
    }
  }

  @Test
  void immutableLookupDoesNotReopenTheDeletedSourceAndRootReplacementRefuses() throws Exception {
    Path directory = Files.createDirectory(root.resolve("work"));
    Path source = Files.writeString(directory.resolve("input"), "one");
    RequestPathIdentity identity = new RequestPathIdentity(directory);
    assertEquals(source, identity.resolve("input"));
    Files.delete(source);
    assertEquals(source, identity.previouslyResolved("input"));
    Files.move(directory, root.resolve("original"));
    Files.createDirectory(directory);
    assertThrows(UnsafePathAccessException.class, () -> identity.resolve("input"));
  }
}
