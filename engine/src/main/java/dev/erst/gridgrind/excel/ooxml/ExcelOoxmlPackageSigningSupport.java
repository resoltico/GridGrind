package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.InvalidSigningConfigurationException;
import dev.erst.gridgrind.excel.WorkbookSecurityException;
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
import javax.xml.crypto.MarshalException;
import javax.xml.crypto.dsig.XMLSignatureException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.dsig.SignatureConfig;
import org.apache.poi.poifs.crypt.dsig.SignatureInfo;
import org.jspecify.annotations.Nullable;

/** Loads signing material and applies validated OOXML package signatures. */
public final class ExcelOoxmlPackageSigningSupport {
  private ExcelOoxmlPackageSigningSupport() {}

  /** Signs one materialized workbook package with the supplied PKCS#12 credentials. */
  public static void signWorkbook(Path workbookPath, ExcelOoxmlSignatureOptions signatureOptions)
      throws IOException {
    SigningMaterial signingMaterial = signingMaterial(signatureOptions);
    try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ_WRITE)) {
      SignatureConfig signatureConfig = new SignatureConfig();
      signatureConfig.setKey(signingMaterial.privateKey());
      signatureConfig.setSigningCertificateChain(signingMaterial.certificateChain());
      signatureConfig.setDigestAlgo(
          ExcelOoxmlSecurityPoiBridge.toPoi(signatureOptions.digestAlgorithm()));
      if (signatureOptions.description() != null) {
        signatureConfig.setSignatureDescription(signatureOptions.description());
      }

      SignatureInfo signatureInfo = ExcelOoxmlDsigProviderSupport.newSignatureInfo();
      signatureInfo.setSignatureConfig(signatureConfig);
      signatureInfo.setOpcPackage(pkg);
      confirmAndVerifySignature(
          () -> {
            signatureInfo.confirmSignature();
            return signatureInfo.verifySignature();
          },
          workbookPath);
    } catch (InvalidFormatException exception) {
      throw new WorkbookSecurityException(
          "Failed to open the OOXML workbook package for signing: " + workbookPath, exception);
    }
  }

  /** Signs and verifies one package through a single checked-exception seam. */
  public static void confirmAndVerifySignature(SignatureWriter signatureWriter, Path workbookPath)
      throws IOException {
    boolean valid;
    try {
      valid = signatureWriter.confirmAndVerify();
    } catch (XMLSignatureException | MarshalException exception) {
      throw new WorkbookSecurityException(
          "Failed to sign the OOXML workbook package: " + workbookPath, exception);
    } catch (RuntimeException exception) {
      throw new WorkbookSecurityException(
          "Unexpected OOXML signing failure for " + workbookPath, exception);
    }
    if (!valid) {
      throw new WorkbookSecurityException(
          "The saved workbook signature did not validate after signing " + workbookPath);
    }
  }

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

  public record SigningMaterial(PrivateKey privateKey, List<X509Certificate> certificateChain) {
    public SigningMaterial {
      Objects.requireNonNull(privateKey, "privateKey must not be null");
      certificateChain = List.copyOf(Objects.requireNonNull(certificateChain, "certificateChain"));
    }
  }

  /** Confirms and verifies one OOXML package signature in a single testable step. */
  @FunctionalInterface
  public interface SignatureWriter {
    /** Signs and immediately verifies the saved package. */
    boolean confirmAndVerify() throws XMLSignatureException, MarshalException;
  }
}
