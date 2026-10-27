package dev.erst.gridgrind.core.model.security;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Resolved credentials remain private while source-free persistence policies are shared. */
class ResolvedPackageSecurityTest {
  @Test
  void sharesCanonicalPoliciesAndRedactsNestedCredentials() {
    var encryption = new OoxmlPersistenceEncryptionSpec.None();
    var signature = new OoxmlPersistenceSignatureSpec.None();
    var plain = new ResolvedPersistenceSecurity(encryption, signature);
    assertSame(encryption, plain.encryption());
    assertSame(signature, plain.signature());
    assertTrue(plain.writesPlaintextUnsigned());
    var signed =
        new ResolvedSignature(
            Path.of("signing.p12"),
            "store-secret",
            "key-secret",
            Optional.of("signer"),
            null,
            Optional.empty());
    var secured =
        new ResolvedPersistenceSecurity(
            new ResolvedPersistenceEncryption.Encrypt(
                new ResolvedEncryption("envelope-secret", null, null)),
            new ResolvedPersistenceSignature.Sign(signed));
    assertFalse(secured.writesPlaintextUnsigned());
    for (String secret : new String[] {"store-secret", "key-secret", "envelope-secret"}) {
      assertFalse(secured.toString().contains(secret));
    }
    assertFalse(
        new ResolvedOpenSecurity.Encrypted("open-secret").toString().contains("open-secret"));
    assertEquals(Optional.of("signer"), signed.alias());
    assertEquals(Optional.empty(), signed.description());
  }

  @Test
  void rejectsBlankResolvedCredentialsAndSigningMetadata() {
    assertThrows(IllegalArgumentException.class, () -> new ResolvedOpenSecurity.Encrypted(" "));
    assertThrows(IllegalArgumentException.class, () -> new ResolvedEncryption(" ", null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignature(
                Path.of("key.p12"), "store", " ", Optional.empty(), null, Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignature(
                Path.of("key.p12"), "store", null, Optional.of(" "), null, Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignature(
                Path.of("key.p12"), "store", null, Optional.empty(), null, Optional.of(" ")));
  }
}
