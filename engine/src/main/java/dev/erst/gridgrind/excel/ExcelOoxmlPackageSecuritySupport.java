package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
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
import java.util.Locale;
import java.util.Objects;
import javax.xml.crypto.MarshalException;
import javax.xml.crypto.dsig.XMLSignatureException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.crypt.dsig.SignatureConfig;
import org.apache.poi.poifs.crypt.dsig.SignatureInfo;
import org.apache.poi.poifs.crypt.dsig.SignaturePart;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * OOXML package-security bridge for encrypted open/save and package-signature inspection/signing.
 */
public final class ExcelOoxmlPackageSecuritySupport {
  private ExcelOoxmlPackageSecuritySupport() {}

  /** Opens one workbook through the full formula runtime plus any required package security. */
  public static ExcelWorkbook openWorkbook(
      Path workbookPath,
      ExcelFormulaEnvironment formulaEnvironment,
      ExcelOoxmlOpenOptions openOptions,
      TempFileFactory tempFileFactory)
      throws IOException {
    try (ReadableWorkbook materialized =
        materializeReadableWorkbook(workbookPath, openOptions, tempFileFactory)) {
      return ExcelWorkbook.openMaterializedWorkbook(
          materialized.workbookPath(),
          formulaEnvironment,
          workbookPath.toAbsolutePath().normalize(),
          materialized.packageSecurity(),
          materialized.sourceEncryptionPassword());
    }
  }

  /**
   * Opens one workbook without an explicit formula environment but with package security support.
   */
  public static ExcelWorkbook openWorkbook(
      Path workbookPath, ExcelOoxmlOpenOptions openOptions, TempFileFactory tempFileFactory)
      throws IOException {
    try (ReadableWorkbook materialized =
        materializeReadableWorkbook(workbookPath, openOptions, tempFileFactory)) {
      return ExcelWorkbook.openMaterializedWorkbook(
          materialized.workbookPath(),
          workbookPath.toAbsolutePath().normalize(),
          materialized.packageSecurity(),
          materialized.sourceEncryptionPassword());
    }
  }

  /** Materializes one readable plain `.xlsx` path from a plain or encrypted source workbook. */
  public static ReadableWorkbook materializeReadableWorkbook(
      Path workbookPath, ExcelOoxmlOpenOptions openOptions, TempFileFactory tempFileFactory)
      throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    ExcelOoxmlOpenOptions effectiveOpenOptions = normalizeOpenOptions(openOptions);

    Path absolutePath = workbookPath.toAbsolutePath().normalize();
    if (!Files.exists(absolutePath)) {
      throw new WorkbookNotFoundException(absolutePath);
    }

    FileMagic fileMagic = fileMagic(absolutePath);
    return switch (fileMagic) {
      case OOXML ->
          new ReadableWorkbook(
              absolutePath,
              inspectPackageSecurity(absolutePath, ExcelOoxmlEncryptionSnapshot.none()),
              null,
              false);
      case OLE2 -> decryptWorkbook(absolutePath, effectiveOpenOptions, tempFileFactory);
      default ->
          throw new IllegalArgumentException(
              "Only .xlsx workbooks are supported; unsupported package magic at " + absolutePath);
    };
  }

  static void saveWorkbook(
      ExcelWorkbook workbook,
      Path targetPath,
      ExcelOoxmlPersistenceOptions persistenceOptions,
      TempFileFactory tempFileFactory)
      throws IOException {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(targetPath, "targetPath must not be null");
    Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");

    Path normalizedTarget = targetPath.toAbsolutePath().normalize();
    Files.createDirectories(
        Objects.requireNonNullElse(normalizedTarget.getParent(), normalizedTarget.getRoot()));

    ExcelOoxmlPersistenceOptions explicitOptions =
        persistenceOptions == null
            ? new ExcelOoxmlPersistenceOptions(null, null)
            : persistenceOptions;

    if (passThroughEligible(workbook, explicitOptions)) {
      copySourceWorkbook(workbook.sourcePath(), normalizedTarget);
      return;
    }

    if (requiresResigning(workbook, explicitOptions)) {
      throw new IllegalArgumentException(
          "The workbook was opened from a signed OOXML package and has been mutated or"
              + " re-emitted. Supply persistence.security.signature to sign the saved output"
              + " instead of silently dropping signatures.");
    }

    ExcelOoxmlPersistenceOptions effectiveOptions = effectiveOptions(workbook, explicitOptions);
    if (effectiveOptions.isEmpty()) {
      workbook.savePlainWorkbook(normalizedTarget);
      return;
    }

    Path plainWorkbookPath = tempFileFactory.createTempFile("gridgrind-ooxml-security-", ".xlsx");
    try {
      workbook.savePlainWorkbook(plainWorkbookPath);
      persistMaterializedWorkbook(
          plainWorkbookPath,
          normalizedTarget,
          workbook.loadedPackageSecurity(),
          workbook.sourceEncryptionPassword(),
          workbook.wasMutatedSinceOpen(),
          effectiveOptions);
    } finally {
      deleteIfExists(plainWorkbookPath);
    }
  }

  /**
   * Persists one materialized plain workbook with the requested encryption and signing settings.
   */
  public static void persistMaterializedWorkbook(
      Path plainWorkbookPath,
      Path targetPath,
      ExcelOoxmlPackageSecuritySnapshot sourceSecurity,
      String sourceEncryptionPassword,
      boolean sourceMutated,
      ExcelOoxmlPersistenceOptions persistenceOptions)
      throws IOException {
    Objects.requireNonNull(plainWorkbookPath, "plainWorkbookPath must not be null");
    Objects.requireNonNull(targetPath, "targetPath must not be null");
    Objects.requireNonNull(sourceSecurity, "sourceSecurity must not be null");
    Objects.requireNonNull(persistenceOptions, "persistenceOptions must not be null");

    if (!sourceSecurity.signatures().isEmpty()
        && sourceMutated
        && persistenceOptions.signature() == null) {
      throw new IllegalArgumentException(
          "The workbook was opened from a signed OOXML package and has been rewritten."
              + " Supply persistence.security.signature to sign the saved output instead of"
              + " silently dropping signatures.");
    }

    ExcelOoxmlPersistenceOptions effectiveOptions =
        effectiveOptions(sourceSecurity, sourceEncryptionPassword, persistenceOptions);
    if (effectiveOptions.signature() != null) {
      signWorkbook(plainWorkbookPath, effectiveOptions.signature());
    }
    if (effectiveOptions.encryption() != null) {
      encryptWorkbook(plainWorkbookPath, targetPath, effectiveOptions.encryption());
    } else {
      Files.move(plainWorkbookPath, targetPath, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
    }
  }

  private static boolean passThroughEligible(
      ExcelWorkbook workbook, ExcelOoxmlPersistenceOptions persistenceOptions) {
    return workbook.sourcePath() != null
        && !workbook.wasMutatedSinceOpen()
        && workbook.loadedPackageSecurity().isSecure()
        && persistenceOptions.isEmpty();
  }

  private static boolean requiresResigning(
      ExcelWorkbook workbook, ExcelOoxmlPersistenceOptions persistenceOptions) {
    return workbook.sourcePath() != null
        && !workbook.loadedPackageSecurity().signatures().isEmpty()
        && persistenceOptions.signature() == null;
  }

  private static ExcelOoxmlPersistenceOptions effectiveOptions(
      ExcelWorkbook workbook, ExcelOoxmlPersistenceOptions persistenceOptions) {
    return effectiveOptions(
        workbook.loadedPackageSecurity(), workbook.sourceEncryptionPassword(), persistenceOptions);
  }

  static ExcelOoxmlPersistenceOptions effectiveOptions(
      ExcelOoxmlPackageSecuritySnapshot sourceSecurity,
      String sourceEncryptionPassword,
      ExcelOoxmlPersistenceOptions persistenceOptions) {
    ExcelOoxmlEncryptionOptions encryption = persistenceOptions.encryption();
    if (encryption == null && sourceSecurity.encryption().encrypted()) {
      if (sourceEncryptionPassword == null) {
        throw new IllegalStateException(
            "Encrypted source workbooks must retain their verified source password while open");
      }
      encryption =
          new ExcelOoxmlEncryptionOptions(
              sourceEncryptionPassword, sourceSecurity.encryption().mode());
    }
    return new ExcelOoxmlPersistenceOptions(encryption, persistenceOptions.signature());
  }

  private static ReadableWorkbook decryptWorkbook(
      Path workbookPath, ExcelOoxmlOpenOptions openOptions, TempFileFactory tempFileFactory)
      throws IOException {
    Path plainWorkbookPath = tempFileFactory.createTempFile("gridgrind-ooxml-decrypted-", ".xlsx");
    boolean success = false;
    try (POIFSFileSystem fileSystem = new POIFSFileSystem(workbookPath.toFile())) {
      if (!isEncryptedOoxmlPackage(fileSystem)) {
        throw new IllegalArgumentException("Only .xlsx workbooks are supported");
      }
      String password = openPassword(workbookPath, openOptions);
      EncryptionInfo encryptionInfo;
      encryptionInfo = readEncryptionInfo(fileSystem, workbookPath);

      Decryptor decryptor = Decryptor.getInstance(encryptionInfo);
      boolean unlocked = verifyPassword(decryptor::verifyPassword, password, workbookPath);
      if (!unlocked) {
        throw new InvalidWorkbookPasswordException(workbookPath);
      }

      materializeDecryptedWorkbook(
          () -> decryptor.getDataStream(fileSystem), plainWorkbookPath, workbookPath);

      ExcelOoxmlEncryptionSnapshot encryption = encryptionSnapshot(encryptionInfo);
      ReadableWorkbook readableWorkbook =
          new ReadableWorkbook(
              plainWorkbookPath,
              inspectPackageSecurity(plainWorkbookPath, encryption),
              password,
              true);
      success = true;
      return readableWorkbook;
    } finally {
      if (!success) {
        deleteIfExists(plainWorkbookPath);
      }
    }
  }

  private static boolean isEncryptedOoxmlPackage(POIFSFileSystem fileSystem) {
    return fileSystem.getRoot().hasEntryCaseInsensitive(Decryptor.DEFAULT_POIFS_ENTRY);
  }

  private static String openPassword(Path workbookPath, ExcelOoxmlOpenOptions openOptions) {
    return switch (normalizeOpenOptions(openOptions)) {
      case ExcelOoxmlOpenOptions.Unencrypted _ ->
          throw new WorkbookPasswordRequiredException(workbookPath);
      case ExcelOoxmlOpenOptions.Encrypted encrypted -> encrypted.password();
    };
  }

  private static ExcelOoxmlOpenOptions normalizeOpenOptions(ExcelOoxmlOpenOptions openOptions) {
    return openOptions == null ? new ExcelOoxmlOpenOptions.Unencrypted() : openOptions;
  }

  static ExcelOoxmlEncryptionSnapshot encryptionSnapshot(EncryptionInfo encryptionInfo)
      throws IOException {
    try {
      return new ExcelOoxmlEncryptionSnapshot(
          true,
          ExcelOoxmlSecurityPoiBridge.fromPoi(encryptionInfo.getEncryptionMode()),
          encryptionInfo.getHeader().getCipherAlgorithm().name(),
          encryptionInfo.getVerifier().getHashAlgorithm().name(),
          encryptionInfo.getHeader().getChainingMode().name(),
          encryptionInfo.getHeader().getKeySize(),
          encryptionInfo.getHeader().getBlockSize(),
          encryptionInfo.getVerifier().getSpinCount());
    } catch (RuntimeException exception) {
      throw new WorkbookSecurityException("Failed to inspect OOXML encryption metadata", exception);
    }
  }

  static EncryptionInfo readEncryptionInfo(POIFSFileSystem fileSystem, Path workbookPath)
      throws IOException {
    try {
      return new EncryptionInfo(fileSystem);
    } catch (IOException exception) {
      throw new IllegalArgumentException(
          "Only OOXML .xlsx workbooks are supported; the file is not a supported encrypted"
              + " OOXML workbook: "
              + workbookPath,
          exception);
    }
  }

  static boolean verifyPassword(
      PasswordVerifier passwordVerifier, String password, Path workbookPath) throws IOException {
    try {
      return passwordVerifier.verifyPassword(password);
    } catch (GeneralSecurityException exception) {
      throw new WorkbookSecurityException(
          "Failed to verify the encrypted workbook password: " + workbookPath, exception);
    }
  }

  static void materializeDecryptedWorkbook(
      DecryptedWorkbookStreamSupplier decryptedWorkbookStreamSupplier,
      Path plainWorkbookPath,
      Path workbookPath)
      throws IOException {
    try (InputStream decryptedStream = decryptedWorkbookStreamSupplier.open();
        OutputStream outputStream = Files.newOutputStream(plainWorkbookPath)) {
      decryptedStream.transferTo(outputStream);
    } catch (GeneralSecurityException exception) {
      throw new WorkbookSecurityException(
          "Failed to decrypt the OOXML workbook package: " + workbookPath, exception);
    }
  }

  static ExcelOoxmlPackageSecuritySnapshot inspectPackageSecurity(
      Path workbookPath, ExcelOoxmlEncryptionSnapshot encryption) throws IOException {
    try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ)) {
      SignatureInfo signatureInfo = new SignatureInfo();
      signatureInfo.setSignatureConfig(new SignatureConfig());
      signatureInfo.setOpcPackage(pkg);

      List<ExcelOoxmlSignatureSnapshot> signatures = new ArrayList<>();
      for (SignaturePart signaturePart : signatureInfo.getSignatureParts()) {
        signatures.add(signatureSnapshot(signaturePart));
      }
      return new ExcelOoxmlPackageSecuritySnapshot(encryption, List.copyOf(signatures));
    } catch (InvalidFormatException | RuntimeException exception) {
      throw new WorkbookSecurityException(
          "Failed to inspect OOXML package signatures for " + workbookPath, exception);
    }
  }

  static ExcelOoxmlSignatureSnapshot signatureSnapshot(SignaturePart signaturePart) {
    X509Certificate signer = signaturePart.getSigner();
    return new ExcelOoxmlSignatureSnapshot(
        signaturePart.getPackagePart().getPartName().getName(),
        signerSubject(signer),
        signerIssuer(signer),
        signerSerialNumber(signer),
        signaturePart.validate()
            ? dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState.VALID
            : dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState.INVALID);
  }

  static String signerSubject(X509Certificate signer) {
    return signer == null ? null : signer.getSubjectX500Principal().getName();
  }

  static String signerIssuer(X509Certificate signer) {
    return signer == null ? null : signer.getIssuerX500Principal().getName();
  }

  static String signerSerialNumber(X509Certificate signer) {
    return signer == null ? null : signer.getSerialNumber().toString(16).toUpperCase(Locale.ROOT);
  }

  static void signWorkbook(Path workbookPath, ExcelOoxmlSignatureOptions signatureOptions)
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

      SignatureInfo signatureInfo = new SignatureInfo();
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

  static void confirmAndVerifySignature(SignatureWriter signatureWriter, Path workbookPath)
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

  static SigningMaterial signingMaterial(ExcelOoxmlSignatureOptions signatureOptions) {
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

  static KeyStore loadSigningKeyStore(
      Path keyStorePath, ExcelOoxmlSignatureOptions signatureOptions) {
    try {
      KeyStore keyStore = KeyStore.getInstance("PKCS12");
      try (InputStream inputStream = Files.newInputStream(keyStorePath)) {
        keyStore.load(inputStream, signatureOptions.keystorePassword().toCharArray());
      }
      return keyStore;
    } catch (IOException | java.security.GeneralSecurityException exception) {
      throw new InvalidSigningConfigurationException(
          "Failed to load signing material from " + keyStorePath, exception);
    }
  }

  static String resolveSigningAlias(
      KeyStore keyStore, Path keyStorePath, ExcelOoxmlSignatureOptions signatureOptions) {
    try {
      return resolveAlias(keyStore, signatureOptions.alias(), signatureOptions);
    } catch (java.security.GeneralSecurityException exception) {
      throw new InvalidSigningConfigurationException(
          "Failed to inspect signing aliases in " + keyStorePath, exception);
    }
  }

  static PrivateKey signingPrivateKey(
      KeyStore keyStore,
      Path keyStorePath,
      String alias,
      ExcelOoxmlSignatureOptions signatureOptions) {
    Key key;
    try {
      key = keyStore.getKey(alias, signatureOptions.keyPassword().toCharArray());
    } catch (java.security.GeneralSecurityException exception) {
      throw new InvalidSigningConfigurationException(
          "Failed to load the signing private key from " + keyStorePath, exception);
    }
    if (!(key instanceof PrivateKey privateKey)) {
      throw new InvalidSigningConfigurationException(
          "Signing alias does not resolve to a private key: " + alias);
    }
    return privateKey;
  }

  static List<X509Certificate> signingCertificateChain(
      KeyStore keyStore, Path keyStorePath, String alias) {
    Certificate[] certificateChain;
    try {
      certificateChain = keyStore.getCertificateChain(alias);
    } catch (java.security.GeneralSecurityException exception) {
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

  static String resolveAlias(
      KeyStore keyStore, String requestedAlias, ExcelOoxmlSignatureOptions signatureOptions)
      throws java.security.GeneralSecurityException {
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

  static void encryptWorkbook(
      Path plainWorkbookPath, Path targetPath, ExcelOoxmlEncryptionOptions encryptionOptions)
      throws IOException {
    EncryptionInfo encryptionInfo =
        new EncryptionInfo(ExcelOoxmlSecurityPoiBridge.toPoi(encryptionOptions.mode()));
    Encryptor encryptor = encryptionInfo.getEncryptor();
    encryptor.confirmPassword(encryptionOptions.password());

    writeEncryptedWorkbook(
        fileSystem -> encryptor.getDataStream(fileSystem), plainWorkbookPath, targetPath);
  }

  static void writeEncryptedWorkbook(
      EncryptedWorkbookStreamSupplier encryptedWorkbookStreamSupplier,
      Path plainWorkbookPath,
      Path targetPath)
      throws IOException {
    try (POIFSFileSystem fileSystem = new POIFSFileSystem()) {
      try (OutputStream encryptedStream = encryptedWorkbookStreamSupplier.open(fileSystem)) {
        Files.copy(plainWorkbookPath, encryptedStream);
      }
      try (OutputStream outputStream = Files.newOutputStream(targetPath)) {
        fileSystem.writeFilesystem(outputStream);
      }
    } catch (GeneralSecurityException exception) {
      throw new WorkbookSecurityException(
          "Failed to encrypt the saved OOXML workbook package: " + targetPath, exception);
    }
  }

  static void copySourceWorkbook(Path sourcePath, Path targetPath) throws IOException {
    Objects.requireNonNull(sourcePath, "sourcePath must not be null");
    if (sourcePath.equals(targetPath)) {
      return;
    }
    Files.copy(sourcePath, targetPath, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
  }

  private static FileMagic fileMagic(Path workbookPath) throws IOException {
    try (InputStream inputStream = Files.newInputStream(workbookPath)) {
      return FileMagic.valueOf(FileMagic.prepareToCheckMagic(inputStream));
    }
  }

  static void deleteIfExists(Path path) {
    if (path == null) {
      return;
    }
    try {
      Files.deleteIfExists(path);
    } catch (IOException ignored) {
      // Best-effort cleanup for executor-owned temporary files only.
    }
  }

  /** Temporary-file factory shared with executor-owned low-memory and package-security paths. */
  @FunctionalInterface
  public interface TempFileFactory {
    /** Creates one temporary file owned by the caller. */
    Path createTempFile(String prefix, String suffix) throws IOException;
  }

  /** One readable plain `.xlsx` path materialized from a possibly encrypted source workbook. */
  public static final class ReadableWorkbook implements AutoCloseable {
    private final Path workbookPath;
    private final ExcelOoxmlPackageSecuritySnapshot packageSecurity;
    private final String sourceEncryptionPassword;
    private final boolean deleteOnClose;

    private ReadableWorkbook(
        Path workbookPath,
        ExcelOoxmlPackageSecuritySnapshot packageSecurity,
        String sourceEncryptionPassword,
        boolean deleteOnClose) {
      this.workbookPath = Objects.requireNonNull(workbookPath, "workbookPath must not be null");
      this.packageSecurity =
          Objects.requireNonNull(packageSecurity, "packageSecurity must not be null");
      this.sourceEncryptionPassword = sourceEncryptionPassword;
      this.deleteOnClose = deleteOnClose;
    }

    public Path workbookPath() {
      return workbookPath;
    }

    public ExcelOoxmlPackageSecuritySnapshot packageSecurity() {
      return packageSecurity;
    }

    public String sourceEncryptionPassword() {
      return sourceEncryptionPassword;
    }

    @Override
    public void close() {
      if (deleteOnClose) {
        deleteIfExists(workbookPath);
      }
    }
  }

  record SigningMaterial(PrivateKey privateKey, List<X509Certificate> certificateChain) {
    SigningMaterial {
      Objects.requireNonNull(privateKey, "privateKey must not be null");
      certificateChain = List.copyOf(Objects.requireNonNull(certificateChain, "certificateChain"));
    }
  }

  /** Minimal password-verification seam used to isolate checked decryption failures in tests. */
  @FunctionalInterface
  interface PasswordVerifier {
    /** Verifies one candidate workbook password. */
    boolean verifyPassword(String password) throws GeneralSecurityException;
  }

  /** Supplies one decrypted OOXML input stream for materialization. */
  @FunctionalInterface
  interface DecryptedWorkbookStreamSupplier {
    /** Opens the decrypted workbook bytes. */
    InputStream open() throws IOException, GeneralSecurityException;
  }

  /** Confirms and verifies one OOXML package signature in a single testable step. */
  @FunctionalInterface
  interface SignatureWriter {
    /** Signs and immediately verifies the saved package. */
    boolean confirmAndVerify() throws XMLSignatureException, MarshalException;
  }

  /** Supplies one encrypted OOXML output stream backed by a POIFS container. */
  @FunctionalInterface
  interface EncryptedWorkbookStreamSupplier {
    /** Opens the encrypted output stream for one POIFS package. */
    OutputStream open(POIFSFileSystem fileSystem) throws IOException, GeneralSecurityException;
  }
}
