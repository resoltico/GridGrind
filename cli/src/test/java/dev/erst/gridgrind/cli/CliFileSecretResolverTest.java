package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.SecretReference;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies the CLI resolves only one named no-follow provider leaf at a time. */
class CliFileSecretResolverTest {
  @TempDir Path temporaryDirectory;

  @Test
  void resolvesOnlyTheExplicitReferenceFile() throws Exception {
    Files.writeString(temporaryDirectory.resolve("report-password"), "p@ssw0rd");
    Files.writeString(temporaryDirectory.resolve("neighbour"), "must-not-be-read");

    try (CliFileSecretResolver resolver = CliFileSecretResolver.open(temporaryDirectory)) {
      assertArrayEquals(
          "p@ssw0rd".toCharArray(), resolver.resolve(new SecretReference("report-password")));
    }
  }

  @Test
  void rejectsProviderLeafSymlinks() throws Exception {
    Path target = Files.writeString(temporaryDirectory.resolve("target"), "secret");
    Files.createSymbolicLink(temporaryDirectory.resolve("report-password"), target.getFileName());

    try (CliFileSecretResolver resolver = CliFileSecretResolver.open(temporaryDirectory)) {
      assertThrows(
          IllegalArgumentException.class,
          () -> resolver.resolve(new SecretReference("report-password")));
    }
  }

  @Test
  void rejectsProviderRootSymlinks() throws Exception {
    Path provider = Files.createDirectory(temporaryDirectory.resolve("provider"));
    Path alias = temporaryDirectory.resolve("provider-alias");
    Files.createSymbolicLink(alias, provider.getFileName());

    assertThrows(IOException.class, () -> CliFileSecretResolver.open(alias));
  }

  @Test
  void preservesHostReadFailureWithoutLeakingSecretMaterial() {
    IOException cause = new IOException("provider unavailable");
    CliSecretReadException exception = new CliSecretReadException("report-password", cause);

    assertEquals("report-password", exception.referenceId());
    assertEquals(cause, exception.getCause());
  }

  @Test
  void rejectsMalformedUtf8WithoutReturningSecretMaterial() throws Exception {
    Files.write(
        temporaryDirectory.resolve("report-password"), new byte[] {(byte) 0xC3, (byte) 0x28});

    try (CliFileSecretResolver resolver = CliFileSecretResolver.open(temporaryDirectory)) {
      assertThrows(
          IllegalArgumentException.class,
          () -> resolver.resolve(new SecretReference("report-password")));
    }
  }

  @Test
  void rejectsMissingAndOversizedProviderLeaves() throws Exception {
    try (CliFileSecretResolver resolver = CliFileSecretResolver.open(temporaryDirectory)) {
      assertThrows(
          CliSecretReadException.class,
          () -> resolver.resolve(new SecretReference("missing-password")));
    }

    Files.write(temporaryDirectory.resolve("oversized-password"), new byte[65_537]);
    try (CliFileSecretResolver resolver = CliFileSecretResolver.open(temporaryDirectory)) {
      assertThrows(
          IllegalArgumentException.class,
          () -> resolver.resolve(new SecretReference("oversized-password")));
    }
  }

  @Test
  void failsClosedWhenTheProviderFilesystemCannotOfferNoFollowDirectoryAccess() throws IOException {
    AtomicBoolean closed = new AtomicBoolean();
    IOException exception;
    try (DirectoryStream<Path> nonSecureDirectory =
        new DirectoryStream<>() {
          @Override
          public Iterator<Path> iterator() {
            return List.<Path>of().iterator();
          }

          @Override
          public void close() {
            closed.set(true);
          }
        }) {
      exception =
          assertThrows(IOException.class, () -> new CliFileSecretResolver(nonSecureDirectory));
    }

    assertEquals(
        "secret-provider filesystem does not support no-follow directory access",
        exception.getMessage());
    assertTrue(closed.get());
  }

  @Test
  @SuppressWarnings("PMD.CloseResource") // The enclosing stream owns the no-follow descriptor.
  void rejectsAProviderLeafThatEndsBeforeTheDescriptorDeclaredSize() throws Exception {
    Files.writeString(temporaryDirectory.resolve("report-password"), "secret");
    try (DirectoryStream<Path> nativeStream = Files.newDirectoryStream(temporaryDirectory)) {
      SecureDirectoryStream<Path> nativeDirectory = (SecureDirectoryStream<Path>) nativeStream;
      try (CliFileSecretResolver resolver =
          new CliFileSecretResolver(
              new FaultInjectingSecureDirectoryStream(nativeDirectory, true, false))) {
        assertThrows(
            IllegalArgumentException.class,
            () -> resolver.resolve(new SecretReference("report-password")));
      }
    }
  }

  @Test
  @SuppressWarnings("PMD.CloseResource") // The enclosing stream owns the no-follow descriptor.
  void managedBindingsAbsorbOwnedSecretDescriptorCloseFailures() throws Exception {
    Path scratchRoot = Files.createTempDirectory("gridgrind-secret-close-failure-");
    try (DirectoryStream<Path> nativeStream = Files.newDirectoryStream(temporaryDirectory)) {
      SecureDirectoryStream<Path> nativeDirectory = (SecureDirectoryStream<Path>) nativeStream;
      try (CliFileSecretResolver resolver =
          new CliFileSecretResolver(
              new FaultInjectingSecureDirectoryStream(nativeDirectory, false, true))) {
        dev.erst.gridgrind.engine.api.GridGrindRequestInputs inputs =
            new dev.erst.gridgrind.engine.api.GridGrindRequestInputs(
                temporaryDirectory, scratchRoot, CliExecutionGrants.denied(), resolver);

        assertDoesNotThrow(
            () ->
                new CliExecutionBindingsFactory.ManagedRequestInputs(
                        inputs, java.util.Optional.empty(), java.util.Optional.of(resolver))
                    .close());
      }
    }
    assertTrue(Files.notExists(scratchRoot));
  }
}
