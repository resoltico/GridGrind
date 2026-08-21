package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.InvalidSigningConfigurationException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.security.Key;
import java.security.KeyStore;
import java.security.PrivateKey;
import java.security.cert.Certificate;
import java.security.cert.X509Certificate;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Resolves verified PKCS#12 private-key and certificate material for OOXML package signing. */
public final class ExcelOoxmlSigningMaterialSupport {
  private ExcelOoxmlSigningMaterialSupport() {}

  /** Resolves one complete signing bundle from the supplied signature options. */
  public static SigningMaterial signingMaterial(ExcelOoxmlSignatureOptions signatureOptions) {
    Path keyStorePath = signatureOptions.pkcs12Path().toAbsolutePath().normalize();
    if (!Files.exists(keyStorePath)) {
      throw new InvalidSigningConfigurationException(
          "Signing keystore does not exist: " + keyStorePath);
    }
    KeyStore keyStore = loadSigningKeyStore(keyStorePath, signatureOptions);
    String alias = resolveSigningAlias(keyStore, keyStorePath, signatureOptions);
    PrivateKey privateKey = signingPrivateKey(keyStore, keyStorePath, alias, signatureOptions);
    List<X509Certificate> certificateChain = signingCertificateChain(keyStore, keyStorePath, alias);
    return new SigningMaterial(privateKey, certificateChain);
  }

  /** Loads the configured PKCS#12 keystore from disk. */
  public static KeyStore loadSigningKeyStore(
      Path keyStorePath, ExcelOoxmlSignatureOptions signatureOptions) {
    try {
      KeyStore keyStore = KeyStore.getInstance("PKCS12");
      try (java.io.InputStream inputStream = Files.newInputStream(keyStorePath)) {
        keyStore.load(inputStream, signatureOptions.keystorePassword().toCharArray());
      }
      return keyStore;
    } catch (IOException | GeneralSecurityException exception) {
      throw new InvalidSigningConfigurationException(
          "Failed to load signing material from " + keyStorePath, exception);
    }
  }

  /** Resolves the concrete signing alias that will be used from one keystore. */
  public static String resolveSigningAlias(
      KeyStore keyStore, Path keyStorePath, ExcelOoxmlSignatureOptions signatureOptions) {
    try {
      return resolveAlias(keyStore, signatureOptions.alias(), signatureOptions);
    } catch (GeneralSecurityException exception) {
      throw new InvalidSigningConfigurationException(
          "Failed to inspect signing aliases in " + keyStorePath, exception);
    }
  }

  /** Reads the private signing key for one resolved alias. */
  public static PrivateKey signingPrivateKey(
      KeyStore keyStore,
      Path keyStorePath,
      String alias,
      ExcelOoxmlSignatureOptions signatureOptions) {
    Key key;
    try {
      key = keyStore.getKey(alias, signatureOptions.keyPassword().toCharArray());
    } catch (GeneralSecurityException exception) {
      throw new InvalidSigningConfigurationException(
          "Failed to load the signing private key from " + keyStorePath, exception);
    }
    if (!(key instanceof PrivateKey privateKey)) {
      throw new InvalidSigningConfigurationException(
          "Signing alias does not resolve to a private key: " + alias);
    }
    return privateKey;
  }

  /** Reads the X.509 certificate chain for one resolved signing alias. */
  public static List<X509Certificate> signingCertificateChain(
      KeyStore keyStore, Path keyStorePath, String alias) {
    Certificate[] certificateChain;
    try {
      certificateChain = keyStore.getCertificateChain(alias);
    } catch (GeneralSecurityException exception) {
      throw new InvalidSigningConfigurationException(
          "Failed to load the signing certificate chain from " + keyStorePath, exception);
    }
    if (certificateChain == null || certificateChain.length == 0) {
      throw new InvalidSigningConfigurationException(
          "Signing alias does not contain an X.509 certificate chain: " + alias);
    }

    List<X509Certificate> x509Chain = new ArrayList<>(certificateChain.length);
    for (Certificate certificate : certificateChain) {
      if (!(certificate instanceof X509Certificate x509Certificate)) {
        throw new InvalidSigningConfigurationException(
            "Signing alias contains a non-X.509 certificate: " + alias);
      }
      x509Chain.add(x509Certificate);
    }
    return List.copyOf(x509Chain);
  }

  /** Resolves the requested alias or the only usable private-key alias in the keystore. */
  public static String resolveAlias(
      KeyStore keyStore,
      @Nullable String requestedAlias,
      ExcelOoxmlSignatureOptions signatureOptions)
      throws GeneralSecurityException {
    if (requestedAlias != null) {
      if (!keyStore.containsAlias(requestedAlias)) {
        throw new InvalidSigningConfigurationException(
            "Signing alias does not exist in the PKCS#12 keystore: " + requestedAlias);
      }
      return requestedAlias;
    }

    List<String> aliases = new ArrayList<>();
    Enumeration<String> aliasEnumeration = keyStore.aliases();
    while (aliasEnumeration.hasMoreElements()) {
      String alias = aliasEnumeration.nextElement();
      if (keyStore.isKeyEntry(alias)
          && keyStore.getKey(alias, signatureOptions.keyPassword().toCharArray())
              instanceof PrivateKey) {
        aliases.add(alias);
      }
    }

    if (aliases.isEmpty()) {
      throw new InvalidSigningConfigurationException(
          "The PKCS#12 keystore does not contain a private-key entry that can sign OOXML packages");
    }
    if (aliases.size() > 1) {
      throw new InvalidSigningConfigurationException(
          "The PKCS#12 keystore contains multiple private-key aliases."
              + " Supply persistence.security.signature.alias explicitly.");
    }
    return aliases.getFirst();
  }

  /** Verified private-key and certificate material ready for a package-signing operation. */
  public record SigningMaterial(PrivateKey privateKey, List<X509Certificate> certificateChain) {
    public SigningMaterial {
      Objects.requireNonNull(privateKey, "privateKey must not be null");
      certificateChain = List.copyOf(Objects.requireNonNull(certificateChain, "certificateChain"));
    }
  }
}
