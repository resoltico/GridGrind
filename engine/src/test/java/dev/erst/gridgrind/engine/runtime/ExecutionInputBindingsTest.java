package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.SecretReference;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Coverage for explicit execution bindings after ambient runtime defaults were removed. */
class ExecutionInputBindingsTest {
  @Test
  void bindingsNormalizeAndExposeTheirExplicitTempRoot() {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            Path.of("tmp", "..", "tmp", "request-root"),
            Path.of("tmp", "..", "tmp", "temp-root"),
            ExecutionGrantTestSupport.noPublication());

    assertTrue(bindings.workingDirectory().isAbsolute());
    assertEquals(Path.of("tmp", "temp-root").toAbsolutePath().normalize(), bindings.tempRoot());
  }

  @Test
  void tempFileFactoryCreatesArtifactsUnderTheExplicitTempRoot() throws IOException {
    Path tempRoot = Files.createTempDirectory("gridgrind-bindings-temp-root-");
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            Files.createTempDirectory("gridgrind-bindings-root-"),
            tempRoot,
            ExecutionGrantTestSupport.noPublication());

    Path created = bindings.tempFileFactory().createTempFile("binding-", ".tmp");

    assertEquals(tempRoot.toAbsolutePath().normalize(), created.getParent());
  }

  @Test
  void requestPathAccessMustBePreparedExactlyOnce() throws IOException {
    Path root = Files.createTempDirectory("gridgrind-bindings-root-");
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            root, root.resolve("private"), ExecutionGrantTestSupport.noPublication());
    assertFalse(bindings.hasRequestPathAccess());
    assertThrows(IllegalStateException.class, bindings::requestPathAccess);

    try (RequestPathAccess access =
        new RequestPathAccess(root, bindings.tempFileFactory(), bindings.executionGrant())) {
      ExecutionInputBindings prepared = bindings.withRequestPathAccess(access);
      assertTrue(prepared.hasRequestPathAccess());
      assertEquals(access, prepared.requestPathAccess());
      assertThrows(IllegalStateException.class, () -> prepared.withRequestPathAccess(access));
    }
  }

  @Test
  void standardInputRequiresExplicitHostAuthorityAndUsesDefensiveCopies() throws IOException {
    Path root = Files.createTempDirectory("gridgrind-standard-input-");
    byte[] source = new byte[] {1, 2, 3};
    ExecutionInputBindings permitted =
        new ExecutionInputBindings(
            root,
            root.resolve("scratch"),
            source,
            ExecutionGrantTestSupport.noPublicationWithStandardInput(List.of()));
    source[0] = 9;

    byte[] firstRead = permitted.standardInputBytes().orElseThrow();
    firstRead[1] = 9;
    assertEquals(1, permitted.standardInputBytes().orElseThrow()[0]);
    assertEquals(2, permitted.standardInputBytes().orElseThrow()[1]);
    assertThrows(
        ExecutionAuthorityDeniedException.class,
        () ->
            new ExecutionInputBindings(
                    root,
                    root.resolve("scratch-denied"),
                    new byte[] {1},
                    ExecutionGrantTestSupport.noPublication())
                .standardInputBytes());
    assertTrue(
        new ExecutionInputBindings(
                root, root.resolve("scratch-empty"), ExecutionGrantTestSupport.noPublication())
            .standardInputBytes()
            .isEmpty());
    assertEquals(
        4,
        new ExecutionInputBindings(
                root,
                root.resolve("scratch-binding"),
                new ExecutionInputBindings.StandardInputBinding(new byte[] {4}),
                ExecutionGrantTestSupport.noPublicationWithStandardInput(List.of()))
            .standardInputBytes()
            .orElseThrow()[0]);
  }

  @Test
  void secretResolutionRequiresBothGrantPermissionAndAHostResolver() throws IOException {
    Path root = Files.createTempDirectory("gridgrind-secret-binding-");
    SecretReference reference = new SecretReference("signing-password");
    char[] material = "s3cret".toCharArray();
    GridGrindExecutionGrant.Bounded permittedGrant =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of(),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.None(),
            List.of(reference),
            GridGrindHostAcceptancePolicy.minimum());
    ExecutionInputBindings permitted =
        new ExecutionInputBindings(
            root, root.resolve("scratch"), permittedGrant, requested -> material);
    ExecutionInputBindings standardInputAndSecret =
        new ExecutionInputBindings(
            root,
            root.resolve("scratch-standard-input"),
            new byte[] {7},
            permittedGrantWithStandardInput(reference),
            requested -> "password".toCharArray());

    assertEquals("s3cret", permitted.resolveSecret(reference));
    assertEquals('\0', material[0]);
    assertEquals(7, standardInputAndSecret.standardInputBytes().orElseThrow()[0]);
    assertThrows(
        ExecutionAuthorityDeniedException.class,
        () ->
            new ExecutionInputBindings(root, root.resolve("scratch"), permittedGrant)
                .resolveSecret(reference));
    assertThrows(
        ExecutionAuthorityDeniedException.class,
        () ->
            new ExecutionInputBindings(
                    root,
                    root.resolve("scratch"),
                    ExecutionGrantTestSupport.noPublication(),
                    requested -> "unused".toCharArray())
                .resolveSecret(reference));
  }

  private static GridGrindExecutionGrant.Bounded permittedGrantWithStandardInput(
      SecretReference reference) {
    return new GridGrindExecutionGrant.Bounded(
        List.of(new GridGrindExecutionGrant.ReadAuthority.StandardInput()),
        List.of(),
        new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        new GridGrindExecutionGrant.PublicationAuthority.None(),
        List.of(reference),
        GridGrindHostAcceptancePolicy.minimum());
  }
}
