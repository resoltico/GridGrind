package dev.erst.gridgrind.engine.security;

import dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment;
import dev.erst.gridgrind.core.model.security.OoxmlEncryptionReport;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec;
import dev.erst.gridgrind.core.model.security.OoxmlSignatureReport;
import dev.erst.gridgrind.core.model.security.ResolvedEncryption;
import dev.erst.gridgrind.core.model.security.ResolvedOpenSecurity;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceEncryption;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceSecurity;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceSignature;
import dev.erst.gridgrind.core.model.security.ResolvedSignature;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteHash;
import dev.erst.gridgrind.engine.backend.xssf.InvalidSigningConfigurationException;
import dev.erst.gridgrind.engine.backend.xssf.InvalidWorkbookPasswordException;
import dev.erst.gridgrind.engine.backend.xssf.UnsupportedSourceEncryptionPreservationException;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookNotFoundException;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookNotOpenableException;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookPasswordRequiredException;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookSecurityException;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import dev.erst.gridgrind.engine.packaging.WorkbookFiles;
import dev.erst.gridgrind.engine.packaging.WorkbookNormalization;
import dev.erst.gridgrind.engine.poi.EnumMappings;
import dev.erst.gridgrind.engine.security.spi.OoxmlRelationshipTransformService;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.security.Key;
import java.security.KeyStore;
import java.security.PrivateKey;
import java.security.Provider;
import java.security.cert.Certificate;
import java.security.cert.X509Certificate;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import javax.xml.crypto.MarshalException;
import javax.xml.crypto.dsig.XMLSignatureException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.poifs.crypt.ChainingMode;
import org.apache.poi.poifs.crypt.CipherAlgorithm;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.apache.poi.poifs.crypt.dsig.SignatureConfig;
import org.apache.poi.poifs.crypt.dsig.SignatureInfo;
import org.apache.poi.poifs.crypt.dsig.SignaturePart;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.jspecify.annotations.Nullable;

/** Owns OOXML open/save encryption, signature verification/signing and security metadata. */
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired"})
public final class SecurityAdapter {
  private SecurityAdapter() {}

  /** Opens one workbook through the full formula runtime plus any required package security. */
  public static XssfWorkbook openWorkbook(
      Path workbookPath,
      ResolvedFormulaEnvironment formulaEnvironment,
      ResolvedOpenSecurity openOptions,
      WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    try (ReadableWorkbook materialized =
        materializeReadableWorkbook(workbookPath, openOptions, tempFileFactory)) {
      return XssfWorkbooks.openMaterializedWorkbook(
          materialized.workbookPath(),
          formulaEnvironment,
          Optional.of(workbookPath.toAbsolutePath().normalize()),
          materialized.packageSecurity(),
          materialized.sourceEncryptionPassword());
    }
  }

  /**
   * Opens one workbook without an explicit formula environment but with package security support.
   */
  public static XssfWorkbook openWorkbook(
      Path workbookPath, ResolvedOpenSecurity openOptions, WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    try (ReadableWorkbook materialized =
        materializeReadableWorkbook(workbookPath, openOptions, tempFileFactory)) {
      return XssfWorkbooks.openMaterializedWorkbook(
          materialized.workbookPath(),
          Optional.of(workbookPath.toAbsolutePath().normalize()),
          materialized.packageSecurity(),
          materialized.sourceEncryptionPassword());
    }
  }

  /** Materializes one readable plain `.xlsx` path from a plain or encrypted source workbook. */
  public static ReadableWorkbook materializeReadableWorkbook(
      Path workbookPath, ResolvedOpenSecurity openOptions, WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    ResolvedOpenSecurity effectiveOpenOptions = normalizeOpenOptions(openOptions);

    Path absolutePath = workbookPath.toAbsolutePath().normalize();
    if (!Files.exists(absolutePath)) {
      throw new WorkbookNotFoundException(absolutePath);
    }

    FileMagic fileMagic = fileMagic(absolutePath);
    return switch (fileMagic) {
      case OOXML ->
          new ReadableWorkbook(
              absolutePath,
              inspectPackageSecurity(absolutePath, new OoxmlEncryptionReport.None()),
              Optional.empty(),
              Optional.empty());
      case OLE2 -> decryptWorkbook(absolutePath, effectiveOpenOptions);
      default ->
          throw new WorkbookNotOpenableException(
              absolutePath,
              new IllegalArgumentException(
                  "Only .xlsx workbooks are supported; unsupported package magic"));
    };
  }

  public static void saveWorkbook(
      XssfWorkbook workbook,
      Path targetPath,
      WorkbookArtifactWriteDisposition writeDisposition,
      ResolvedPersistenceSecurity persistenceOptions,
      WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(targetPath, "targetPath must not be null");
    Objects.requireNonNull(writeDisposition, "writeDisposition must not be null");
    Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");

    Path normalizedTarget = targetPath.toAbsolutePath().normalize();
    createTargetParentDirectories(normalizedTarget);

    ResolvedPersistenceSecurity explicitOptions =
        persistenceOptions == null ? ResolvedPersistenceSecurity.none() : persistenceOptions;

    ResolvedPersistenceSecurity effectiveOptions =
        effectiveOptions(
            workbook.persistence().loadedPackageSecurity(),
            workbook.persistence().sourceEncryptionPassword(),
            explicitOptions);
    if (effectiveOptions.writesPlaintextUnsigned()) {
      workbook.persistence().savePlainWorkbook(normalizedTarget, writeDisposition);
      removeSignatures(normalizedTarget);
      WorkbookNormalization.normalizeWorkbookPackage(normalizedTarget);
      return;
    }

    try (PrivateWorkbookFile privateWorkbook =
        PrivateWorkbookFile.create("gridgrind-ooxml-security-", ".xlsx")) {
      Path plainWorkbookPath =
          WorkbookTempFileFactory.prepareCreateNewTarget(privateWorkbook.workbookPath());
      workbook
          .persistence()
          .savePlainWorkbook(plainWorkbookPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
      persistMaterializedWorkbook(
          plainWorkbookPath,
          normalizedTarget,
          workbook.persistence().loadedPackageSecurity(),
          workbook.persistence().sourceEncryptionPassword(),
          writeDisposition,
          effectiveOptions);
    }
  }

  /**
   * Persists one materialized plain workbook with the requested encryption and signing settings.
   */
  public static void persistMaterializedWorkbook(
      Path plainWorkbookPath,
      Path targetPath,
      OoxmlPackageSecurityReport sourceSecurity,
      Optional<String> sourceEncryptionPassword,
      WorkbookArtifactWriteDisposition writeDisposition,
      ResolvedPersistenceSecurity persistenceOptions)
      throws IOException {
    Objects.requireNonNull(plainWorkbookPath, "plainWorkbookPath must not be null");
    Objects.requireNonNull(targetPath, "targetPath must not be null");
    Objects.requireNonNull(sourceSecurity, "sourceSecurity must not be null");
    Objects.requireNonNull(writeDisposition, "writeDisposition must not be null");
    Objects.requireNonNull(persistenceOptions, "persistenceOptions must not be null");

    Path normalizedTarget = targetPath.toAbsolutePath().normalize();
    createTargetParentDirectories(normalizedTarget);

    ResolvedPersistenceSecurity effectiveOptions =
        effectiveOptions(sourceSecurity, sourceEncryptionPassword, persistenceOptions);
    removeSignatures(plainWorkbookPath);
    WorkbookNormalization.normalizeWorkbookPackage(plainWorkbookPath);
    if (effectiveOptions.signature() instanceof ResolvedPersistenceSignature.Sign sign) {
      signWorkbook(plainWorkbookPath, sign.options());
    }
    if (effectiveOptions.encryption() instanceof ResolvedPersistenceEncryption.Encrypt encrypt) {
      encryptWorkbook(plainWorkbookPath, normalizedTarget, writeDisposition, encrypt.options());
    } else {
      WorkbookFiles.moveWorkbook(plainWorkbookPath, normalizedTarget, writeDisposition);
    }
  }

  /** One readable plain `.xlsx` path materialized from a possibly encrypted source workbook. */
  public static final class ReadableWorkbook implements AutoCloseable {
    private final Path workbookPath;
    private final OoxmlPackageSecurityReport packageSecurity;
    private final Optional<String> sourceEncryptionPassword;
    private final Optional<Path> cleanupRoot;

    ReadableWorkbook(
        Path workbookPath,
        OoxmlPackageSecurityReport packageSecurity,
        Optional<String> sourceEncryptionPassword,
        Optional<Path> cleanupRoot) {
      this.workbookPath = Objects.requireNonNull(workbookPath, "workbookPath must not be null");
      this.packageSecurity =
          Objects.requireNonNull(packageSecurity, "packageSecurity must not be null");
      this.sourceEncryptionPassword =
          Objects.requireNonNull(
              sourceEncryptionPassword, "sourceEncryptionPassword must not be null");
      this.cleanupRoot = Objects.requireNonNull(cleanupRoot, "cleanupRoot must not be null");
    }

    public Path workbookPath() {
      return workbookPath;
    }

    public OoxmlPackageSecurityReport packageSecurity() {
      return packageSecurity;
    }

    public Optional<String> sourceEncryptionPassword() {
      return sourceEncryptionPassword;
    }

    @Override
    public void close() {
      if (cleanupRoot.isPresent()) {
        WorkbookFiles.deleteTreeIfExists(cleanupRoot.orElseThrow());
      }
    }
  }

  /** Reads one encrypted workbook's factual OOXML encryption metadata. */
  public static OoxmlEncryptionReport encryptionSnapshot(EncryptionInfo encryptionInfo)
      throws IOException {
    try {
      return new OoxmlEncryptionReport.Encrypted(
          fromPoi(encryptionInfo.getEncryptionMode()),
          fromPoi(encryptionInfo.getHeader().getCipherAlgorithm()),
          fromPoi(encryptionInfo.getVerifier().getHashAlgorithm()),
          fromPoi(encryptionInfo.getHeader().getChainingMode()),
          encryptionInfo.getHeader().getKeySize(),
          encryptionInfo.getHeader().getBlockSize(),
          encryptionInfo.getVerifier().getSpinCount());
    } catch (RuntimeException exception) {
      throw new WorkbookSecurityException("Failed to inspect OOXML encryption metadata", exception);
    }
  }

  /** Loads POI's encryption envelope from one encrypted OOXML package container. */
  public static EncryptionInfo readEncryptionInfo(POIFSFileSystem fileSystem, Path workbookPath)
      throws IOException {
    try {
      return new EncryptionInfo(fileSystem);
    } catch (IOException exception) {
      throw new WorkbookNotOpenableException(workbookPath, exception);
    }
  }

  /** Verifies one candidate workbook password through a checked-exception seam. */
  public static boolean verifyPassword(
      PasswordVerifier passwordVerifier, String password, Path workbookPath) throws IOException {
    try {
      return passwordVerifier.verifyPassword(password);
    } catch (GeneralSecurityException exception) {
      throw new WorkbookSecurityException(
          "Failed to verify the encrypted workbook password: " + workbookPath, exception);
    }
  }

  /** Materializes decrypted OOXML bytes onto one temporary plain workbook path. */
  public static void materializeDecryptedWorkbook(
      DecryptedWorkbookStreamSupplier decryptedWorkbookStreamSupplier,
      Path plainWorkbookPath,
      Path workbookPath)
      throws IOException {
    try (java.io.InputStream decryptedStream = decryptedWorkbookStreamSupplier.open();
        java.io.OutputStream outputStream = Files.newOutputStream(plainWorkbookPath)) {
      decryptedStream.transferTo(outputStream);
    } catch (GeneralSecurityException exception) {
      throw new WorkbookSecurityException(
          "Failed to decrypt the OOXML workbook package: " + workbookPath, exception);
    }
  }

  /** Encrypts one plain materialized workbook into the requested OOXML encryption mode. */
  public static void encryptWorkbook(
      Path plainWorkbookPath,
      Path targetPath,
      WorkbookArtifactWriteDisposition writeDisposition,
      ResolvedEncryption encryptionOptions)
      throws IOException {
    CipherAlgorithm cipherAlgorithm = toPoi(encryptionOptions.cipher());
    EncryptionInfo encryptionInfo =
        new EncryptionInfo(
            EncryptionMode.agile,
            cipherAlgorithm,
            toPoi(encryptionOptions.hash()),
            cipherAlgorithm.defaultKeySize,
            cipherAlgorithm.blockSize,
            toPoi(dev.erst.gridgrind.core.vocabulary.ExcelOoxmlChainingMode.CBC));
    Encryptor encryptor = encryptionInfo.getEncryptor();
    encryptor.confirmPassword(encryptionOptions.password());

    writeEncryptedWorkbook(
        fileSystem -> encryptor.getDataStream(fileSystem),
        plainWorkbookPath,
        targetPath,
        writeDisposition);
  }

  /** Writes one encrypted POIFS-backed OOXML package to the target workbook path. */
  public static void writeEncryptedWorkbook(
      EncryptedWorkbookStreamSupplier encryptedWorkbookStreamSupplier,
      Path plainWorkbookPath,
      Path targetPath,
      WorkbookArtifactWriteDisposition writeDisposition)
      throws IOException {
    try (POIFSFileSystem fileSystem = new POIFSFileSystem()) {
      try (java.io.OutputStream encryptedStream =
          encryptedWorkbookStreamSupplier.open(fileSystem)) {
        Files.copy(plainWorkbookPath, encryptedStream);
      }
      try (java.io.OutputStream outputStream =
          WorkbookFiles.newWorkbookOutputStream(targetPath, writeDisposition)) {
        fileSystem.writeFilesystem(outputStream);
      }
    } catch (GeneralSecurityException exception) {
      throw new WorkbookSecurityException(
          "Failed to encrypt the saved OOXML workbook package: " + targetPath, exception);
    }
  }

  /** Minimal password-verification seam used to isolate checked decryption failures in tests. */
  @FunctionalInterface
  public interface PasswordVerifier {
    /** Verifies one candidate workbook password. */
    boolean verifyPassword(String password) throws GeneralSecurityException;
  }

  /** Supplies one decrypted OOXML input stream for materialization. */
  @FunctionalInterface
  public interface DecryptedWorkbookStreamSupplier {
    /** Opens the decrypted workbook bytes. */
    java.io.InputStream open() throws IOException, GeneralSecurityException;
  }

  /** Supplies one encrypted OOXML output stream backed by a POIFS container. */
  @FunctionalInterface
  public interface EncryptedWorkbookStreamSupplier {
    /** Opens the encrypted output stream for one POIFS package. */
    java.io.OutputStream open(POIFSFileSystem fileSystem)
        throws IOException, GeneralSecurityException;
  }

  /** Computes the effective encryption and signing settings for one persistence request. */
  public static ResolvedPersistenceSecurity effectiveOptions(
      OoxmlPackageSecurityReport sourceSecurity,
      java.util.Optional<String> sourceEncryptionPassword,
      ResolvedPersistenceSecurity persistenceOptions) {
    return switch (persistenceOptions.encryption()) {
      case OoxmlPersistenceEncryptionSpec.None _ -> persistenceOptions;
      case ResolvedPersistenceEncryption.Encrypt _ -> persistenceOptions;
      case OoxmlPersistenceEncryptionSpec.PreserveSource _ ->
          new ResolvedPersistenceSecurity(
              new ResolvedPersistenceEncryption.Encrypt(
                  preservedSourceEncryption(sourceSecurity, sourceEncryptionPassword)),
              persistenceOptions.signature());
    };
  }

  private static ResolvedEncryption preservedSourceEncryption(
      OoxmlPackageSecurityReport sourceSecurity,
      java.util.Optional<String> sourceEncryptionPassword) { // LIM-038
    if (!(sourceSecurity.encryption()
        instanceof OoxmlEncryptionReport.Encrypted sourceEncryption)) {
      throw new IllegalArgumentException(
          "PRESERVE_SOURCE encryption requires an encrypted source workbook");
    }
    if (sourceEncryptionPassword.isEmpty()) {
      throw new IllegalStateException(
          "Encrypted source workbooks must retain their verified source password while open");
    }
    if (sourceEncryption.mode()
        != dev.erst.gridgrind.core.vocabulary.ExcelOoxmlEncryptionMode.AGILE) {
      throw new UnsupportedSourceEncryptionPreservationException(
          "Source workbook encryption mode "
              + sourceEncryption.mode()
              + " is readable but not auto-preservable on write;"
              + " set persistence.security.encryption explicitly to author a supported AGILE"
              + " package instead.");
    }
    if (sourceEncryption.chainingMode()
        != dev.erst.gridgrind.core.vocabulary.ExcelOoxmlChainingMode.CBC) {
      throw new UnsupportedSourceEncryptionPreservationException(
          "Source workbook encryption chaining mode "
              + sourceEncryption.chainingMode()
              + " is readable but not auto-preservable on write;"
              + " set persistence.security.encryption explicitly to author a supported AGILE"
              + " package instead.");
    }
    return new ResolvedEncryption(
        sourceEncryptionPassword.orElseThrow(),
        toWriteCipher(sourceEncryption.cipherAlgorithm())
            .orElseThrow(
                () ->
                    new UnsupportedSourceEncryptionPreservationException(
                        "Source workbook encryption cipher "
                            + sourceEncryption.cipherAlgorithm()
                            + " is readable but not auto-preservable on write;"
                            + " set persistence.security.encryption explicitly to author a"
                            + " supported AGILE package instead.")),
        toWriteHash(sourceEncryption.hashAlgorithm())
            .orElseThrow(
                () ->
                    new UnsupportedSourceEncryptionPreservationException(
                        "Source workbook encryption hash "
                            + sourceEncryption.hashAlgorithm()
                            + " is readable but not auto-preservable on write;"
                            + " set persistence.security.encryption explicitly to author a"
                            + " supported AGILE package instead.")));
  }

  /** Reads one workbook package's signature state alongside its already-known encryption facts. */
  public static OoxmlPackageSecurityReport inspectPackageSecurity(
      Path workbookPath, OoxmlEncryptionReport encryption) throws IOException {
    OPCPackage pkg;
    try {
      pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ);
    } catch (InvalidFormatException | NotOfficeXmlFileException exception) {
      throw new WorkbookNotOpenableException(workbookPath, exception);
    }
    try (pkg) {
      return inspectOpenedPackage(pkg, workbookPath, encryption);
    }
  }

  private static OoxmlPackageSecurityReport inspectOpenedPackage(
      OPCPackage pkg, Path workbookPath, OoxmlEncryptionReport encryption) throws IOException {
    return inspectOpenedPackage(pkg, workbookPath, encryption, SecurityAdapter::newSignatureInfo);
  }

  static OoxmlPackageSecurityReport inspectOpenedPackage(
      OPCPackage pkg,
      Path workbookPath,
      OoxmlEncryptionReport encryption,
      SignatureInfoFactory signatureInfoFactory)
      throws IOException {
    try {
      SignatureInfo signatureInfo = signatureInfoFactory.create();
      signatureInfo.setSignatureConfig(new SignatureConfig());
      signatureInfo.setOpcPackage(pkg);

      List<OoxmlSignatureReport> signatures = new ArrayList<>();
      for (SignaturePart signaturePart : signatureInfo.getSignatureParts()) {
        signatures.add(signatureSnapshot(signaturePart));
      }
      return new OoxmlPackageSecurityReport(encryption, List.copyOf(signatures));
    } catch (RuntimeException exception) {
      throw new WorkbookSecurityException(
          "Failed to inspect OOXML package signatures for " + workbookPath, exception);
    }
  }

  /** Converts one POI signature part into the public GridGrind signature snapshot. */
  public static OoxmlSignatureReport signatureSnapshot(SignaturePart signaturePart) {
    boolean valid = signaturePart.validate();
    X509Certificate signerCertificate = signaturePart.getSigner();
    return new OoxmlSignatureReport(
        signaturePart.getPackagePart().getPartName().getName(),
        signerIdentity(signerCertificate),
        valid ? ExcelOoxmlSignatureState.VALID : ExcelOoxmlSignatureState.INVALID);
  }

  /** Returns the factual signer identity when one certificate is present. */
  public static Optional<OoxmlSignatureReport.SignerIdentity> signerIdentity(
      @Nullable X509Certificate signer) {
    return signer == null
        ? Optional.empty()
        : Optional.of(
            new OoxmlSignatureReport.SignerIdentity(
                signer.getSubjectX500Principal().getName(),
                signer.getIssuerX500Principal().getName(),
                signer.getSerialNumber().toString(16).toUpperCase(Locale.ROOT)));
  }

  /** Creates the POI signature-inspection adapter for an already-opened OOXML package. */
  @FunctionalInterface
  interface SignatureInfoFactory {
    /** Returns one configured signature-inspection adapter. */
    SignatureInfo create();
  }

  /** Signs one materialized workbook package with the supplied PKCS#12 credentials. */
  public static void signWorkbook(Path workbookPath, ResolvedSignature signatureOptions)
      throws IOException {
    SigningMaterial signingMaterial = signingMaterial(signatureOptions);
    try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ_WRITE)) {
      SignatureConfig signatureConfig = new SignatureConfig();
      signatureConfig.setKey(signingMaterial.privateKey());
      signatureConfig.setSigningCertificateChain(signingMaterial.certificateChain());
      signatureConfig.setDigestAlgo(toPoi(signatureOptions.digestAlgorithm()));
      if (signatureOptions.description().isPresent()) {
        signatureConfig.setSignatureDescription(signatureOptions.description().orElseThrow());
      }

      SignatureInfo signatureInfo = newSignatureInfo();
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

  /** Confirms and verifies one OOXML package signature in a single testable step. */
  @FunctionalInterface
  public interface SignatureWriter {
    /** Signs and immediately verifies the saved package. */
    boolean confirmAndVerify() throws XMLSignatureException, MarshalException;
  }

  private static final String DIGITAL_SIGNATURE_ORIGIN_RELATIONSHIP =
      "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
  private static final String SIGNATURE_PART_DIRECTORY = "/_xmlsignatures/";

  public static void removeSignatures(Path workbookPath) throws IOException {
    try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ_WRITE)) {
      List<String> originRelationshipIds = new ArrayList<>();
      for (PackageRelationship relationship : pkg.getRelationships()) {
        if (DIGITAL_SIGNATURE_ORIGIN_RELATIONSHIP.equals(relationship.getRelationshipType())) {
          originRelationshipIds.add(relationship.getId());
        }
      }
      List<PackagePartName> signaturePartNames =
          pkg.getParts().stream()
              .map(part -> part.getPartName())
              .filter(partName -> partName.getName().startsWith(SIGNATURE_PART_DIRECTORY))
              .toList();
      for (PackagePartName signaturePartName : signaturePartNames) {
        pkg.removePart(signaturePartName);
      }
      for (String originRelationshipId : originRelationshipIds) {
        pkg.removeRelationship(originRelationshipId);
      }
    } catch (InvalidFormatException | NotOfficeXmlFileException exception) {
      throw new WorkbookSecurityException(
          "Failed to open the OOXML workbook package for signature removal: " + workbookPath,
          exception);
    }
  }

  static ResolvedOpenSecurity normalizeOpenOptions(ResolvedOpenSecurity openOptions) {
    return openOptions == null ? new ResolvedOpenSecurity.Unencrypted() : openOptions;
  }

  static ReadableWorkbook decryptWorkbook(Path workbookPath, ResolvedOpenSecurity openOptions)
      throws IOException {
    try (PrivateWorkbookFile privateWorkbook =
            PrivateWorkbookFile.create("gridgrind-ooxml-decrypted-", ".xlsx");
        POIFSFileSystem fileSystem = new POIFSFileSystem(workbookPath.toFile())) {
      if (!isEncryptedOoxmlPackage(fileSystem)) {
        throw new WorkbookNotOpenableException(
            workbookPath,
            new IllegalArgumentException("Only encrypted OOXML .xlsx packages are supported"));
      }
      String password = openPassword(openOptions);
      EncryptionInfo encryptionInfo = readEncryptionInfo(fileSystem, workbookPath);

      Decryptor decryptor = Decryptor.getInstance(encryptionInfo);
      boolean unlocked = verifyPassword(decryptor::verifyPassword, password, workbookPath);
      if (!unlocked) {
        throw new InvalidWorkbookPasswordException();
      }

      materializeDecryptedWorkbook(
          () -> decryptor.getDataStream(fileSystem), privateWorkbook.workbookPath(), workbookPath);

      OoxmlEncryptionReport encryption = encryptionSnapshot(encryptionInfo);
      return new ReadableWorkbook(
          privateWorkbook.workbookPath(),
          inspectPackageSecurity(privateWorkbook.workbookPath(), encryption),
          Optional.of(password),
          Optional.of(privateWorkbook.releaseRoot()));
    }
  }

  static void createTargetParentDirectories(Path targetPath) throws IOException {
    Path parentOrRoot = Objects.requireNonNullElse(targetPath.getParent(), targetPath.getRoot());
    Files.createDirectories(parentOrRoot);
  }

  static FileMagic fileMagic(Path workbookPath) throws IOException {
    try (java.io.InputStream inputStream = Files.newInputStream(workbookPath)) {
      return FileMagic.valueOf(FileMagic.prepareToCheckMagic(inputStream));
    }
  }

  private static boolean isEncryptedOoxmlPackage(POIFSFileSystem fileSystem) {
    return fileSystem.getRoot().hasEntryCaseInsensitive(Decryptor.DEFAULT_POIFS_ENTRY);
  }

  private static String openPassword(ResolvedOpenSecurity openOptions) {
    return switch (normalizeOpenOptions(openOptions)) {
      case ResolvedOpenSecurity.Unencrypted _ -> throw new WorkbookPasswordRequiredException();
      case ResolvedOpenSecurity.Encrypted encrypted -> encrypted.password();
    };
  }

  private static final Map<ExcelOoxmlEncryptionMode, EncryptionMode> TO_POI_ENCRYPTION_MODE =
      EnumMappings.exactEnumMap(
          ExcelOoxmlEncryptionMode.class,
          "Apache POI OOXML encryption-mode mapping",
          Map.ofEntries(
              Map.entry(ExcelOoxmlEncryptionMode.AGILE, EncryptionMode.agile),
              Map.entry(ExcelOoxmlEncryptionMode.STANDARD, EncryptionMode.standard)));

  private static final Map<EncryptionMode, ExcelOoxmlEncryptionMode> FROM_POI_ENCRYPTION_MODE =
      Map.ofEntries(
          Map.entry(EncryptionMode.agile, ExcelOoxmlEncryptionMode.AGILE),
          Map.entry(EncryptionMode.standard, ExcelOoxmlEncryptionMode.STANDARD));

  private static final Map<CipherAlgorithm, ExcelOoxmlCipherAlgorithm> FROM_POI_CIPHER_ALGORITHM =
      EnumMappings.exactEnumMap(
          CipherAlgorithm.class,
          "GridGrind OOXML cipher-algorithm mapping",
          Map.ofEntries(
              Map.entry(CipherAlgorithm.rc4, ExcelOoxmlCipherAlgorithm.RC4),
              Map.entry(CipherAlgorithm.aes128, ExcelOoxmlCipherAlgorithm.AES_128),
              Map.entry(CipherAlgorithm.aes192, ExcelOoxmlCipherAlgorithm.AES_192),
              Map.entry(CipherAlgorithm.aes256, ExcelOoxmlCipherAlgorithm.AES_256),
              Map.entry(CipherAlgorithm.rc2, ExcelOoxmlCipherAlgorithm.RC2),
              Map.entry(CipherAlgorithm.des, ExcelOoxmlCipherAlgorithm.DES),
              Map.entry(CipherAlgorithm.des3, ExcelOoxmlCipherAlgorithm.TRIPLE_DES),
              Map.entry(CipherAlgorithm.des3_112, ExcelOoxmlCipherAlgorithm.TRIPLE_DES_112),
              Map.entry(CipherAlgorithm.rsa, ExcelOoxmlCipherAlgorithm.RSA)));

  private static final Map<HashAlgorithm, ExcelOoxmlHashAlgorithm> FROM_POI_HASH_ALGORITHM =
      EnumMappings.exactEnumMap(
          HashAlgorithm.class,
          "GridGrind OOXML hash-algorithm mapping",
          Map.ofEntries(
              Map.entry(HashAlgorithm.none, ExcelOoxmlHashAlgorithm.NONE),
              Map.entry(HashAlgorithm.sha1, ExcelOoxmlHashAlgorithm.SHA_1),
              Map.entry(HashAlgorithm.sha224, ExcelOoxmlHashAlgorithm.SHA_224),
              Map.entry(HashAlgorithm.sha256, ExcelOoxmlHashAlgorithm.SHA_256),
              Map.entry(HashAlgorithm.sha384, ExcelOoxmlHashAlgorithm.SHA_384),
              Map.entry(HashAlgorithm.sha512, ExcelOoxmlHashAlgorithm.SHA_512),
              Map.entry(HashAlgorithm.md2, ExcelOoxmlHashAlgorithm.MD2),
              Map.entry(HashAlgorithm.md4, ExcelOoxmlHashAlgorithm.MD4),
              Map.entry(HashAlgorithm.md5, ExcelOoxmlHashAlgorithm.MD5),
              Map.entry(HashAlgorithm.ripemd128, ExcelOoxmlHashAlgorithm.RIPEMD_128),
              Map.entry(HashAlgorithm.ripemd160, ExcelOoxmlHashAlgorithm.RIPEMD_160),
              Map.entry(HashAlgorithm.ripemd256, ExcelOoxmlHashAlgorithm.RIPEMD_256),
              Map.entry(HashAlgorithm.whirlpool, ExcelOoxmlHashAlgorithm.WHIRLPOOL)));

  private static final Map<ChainingMode, ExcelOoxmlChainingMode> FROM_POI_CHAINING_MODE =
      EnumMappings.exactEnumMap(
          ChainingMode.class,
          "GridGrind OOXML chaining-mode mapping",
          Map.ofEntries(
              Map.entry(ChainingMode.ecb, ExcelOoxmlChainingMode.ECB),
              Map.entry(ChainingMode.cbc, ExcelOoxmlChainingMode.CBC),
              Map.entry(ChainingMode.cfb, ExcelOoxmlChainingMode.CFB)));

  private static final Map<ExcelOoxmlSignatureDigestAlgorithm, HashAlgorithm>
      TO_POI_SIGNATURE_DIGEST_ALGORITHM =
          EnumMappings.exactEnumMap(
              ExcelOoxmlSignatureDigestAlgorithm.class,
              "Apache POI OOXML signature-digest mapping",
              Map.ofEntries(
                  Map.entry(ExcelOoxmlSignatureDigestAlgorithm.SHA256, HashAlgorithm.sha256),
                  Map.entry(ExcelOoxmlSignatureDigestAlgorithm.SHA384, HashAlgorithm.sha384),
                  Map.entry(ExcelOoxmlSignatureDigestAlgorithm.SHA512, HashAlgorithm.sha512)));

  public static EncryptionMode toPoi(ExcelOoxmlEncryptionMode mode) {
    return EnumMappings.requireMappedValue(
        TO_POI_ENCRYPTION_MODE, mode, "GridGrind OOXML encryption mode");
  }

  public static ExcelOoxmlEncryptionMode fromPoi(EncryptionMode mode) {
    return EnumMappings.requireMappedValue(FROM_POI_ENCRYPTION_MODE, mode, "OOXML encryption mode");
  }

  public static CipherAlgorithm toPoi(ExcelOoxmlWriteCipher cipher) {
    Objects.requireNonNull(cipher, "cipher must not be null");
    return switch (cipher) {
      case AES_256 -> CipherAlgorithm.aes256;
      case AES_192 -> CipherAlgorithm.aes192;
    };
  }

  public static ExcelOoxmlCipherAlgorithm fromPoi(CipherAlgorithm algorithm) {
    Objects.requireNonNull(algorithm, "algorithm must not be null");
    return EnumMappings.requireMappedValue(
        FROM_POI_CIPHER_ALGORITHM, algorithm, "OOXML cipher algorithm");
  }

  public static HashAlgorithm toPoi(ExcelOoxmlWriteHash hash) {
    Objects.requireNonNull(hash, "hash must not be null");
    return switch (hash) {
      case SHA_512 -> HashAlgorithm.sha512;
      case SHA_384 -> HashAlgorithm.sha384;
      case SHA_256 -> HashAlgorithm.sha256;
    };
  }

  public static ExcelOoxmlHashAlgorithm fromPoi(HashAlgorithm algorithm) {
    Objects.requireNonNull(algorithm, "algorithm must not be null");
    return EnumMappings.requireMappedValue(
        FROM_POI_HASH_ALGORITHM, algorithm, "OOXML hash algorithm");
  }

  public static ChainingMode toPoi(ExcelOoxmlChainingMode chainingMode) {
    Objects.requireNonNull(chainingMode, "chainingMode must not be null");
    return switch (chainingMode) {
      case ECB -> ChainingMode.ecb;
      case CBC -> ChainingMode.cbc;
      case CFB -> ChainingMode.cfb;
    };
  }

  public static ExcelOoxmlChainingMode fromPoi(ChainingMode chainingMode) {
    Objects.requireNonNull(chainingMode, "chainingMode must not be null");
    return EnumMappings.requireMappedValue(
        FROM_POI_CHAINING_MODE, chainingMode, "OOXML chaining mode");
  }

  public static Optional<ExcelOoxmlWriteCipher> toWriteCipher(ExcelOoxmlCipherAlgorithm cipher) {
    Objects.requireNonNull(cipher, "cipher must not be null");
    if (cipher == ExcelOoxmlCipherAlgorithm.AES_256) {
      return Optional.of(ExcelOoxmlWriteCipher.AES_256);
    }
    if (cipher == ExcelOoxmlCipherAlgorithm.AES_192) {
      return Optional.of(ExcelOoxmlWriteCipher.AES_192);
    }
    return Optional.empty();
  }

  public static Optional<ExcelOoxmlWriteHash> toWriteHash(ExcelOoxmlHashAlgorithm hash) {
    Objects.requireNonNull(hash, "hash must not be null");
    if (hash == ExcelOoxmlHashAlgorithm.SHA_512) {
      return Optional.of(ExcelOoxmlWriteHash.SHA_512);
    }
    if (hash == ExcelOoxmlHashAlgorithm.SHA_384) {
      return Optional.of(ExcelOoxmlWriteHash.SHA_384);
    }
    if (hash == ExcelOoxmlHashAlgorithm.SHA_256) {
      return Optional.of(ExcelOoxmlWriteHash.SHA_256);
    }
    return Optional.empty();
  }

  public static HashAlgorithm toPoi(ExcelOoxmlSignatureDigestAlgorithm algorithm) {
    Objects.requireNonNull(algorithm, "algorithm must not be null");
    return EnumMappings.requireMappedValue(
        TO_POI_SIGNATURE_DIGEST_ALGORITHM, algorithm, "GridGrind OOXML signature digest");
  }

  /** Provider registration is lazy, as it was before feature consolidation. */
  private static final class DsigProvider {
    private static final Provider INSTANCE = newXmlDsigProvider();
  }

  /** Returns a POI signature wrapper bound to GridGrind's xmlsec-4-compatible dsig provider. */
  public static SignatureInfo newSignatureInfo() {
    SignatureInfo signatureInfo = new SignatureInfo();
    signatureInfo.setProvider(provider());
    return signatureInfo;
  }

  public static Provider provider() {
    return DsigProvider.INSTANCE;
  }

  private static Provider newXmlDsigProvider() {
    Provider provider = new org.apache.jcp.xml.dsig.internal.dom.XMLDSigRI();
    provider.put(
        "TransformService." + OoxmlRelationshipTransformService.TRANSFORM_URI,
        OoxmlRelationshipTransformService.class.getName());
    provider.put(
        "TransformService." + OoxmlRelationshipTransformService.TRANSFORM_URI + " MechanismType",
        "DOM");
    return provider;
  }

  /** Resolves one complete signing bundle from the supplied signature options. */
  public static SigningMaterial signingMaterial(ResolvedSignature signatureOptions) {
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
      Path keyStorePath, ResolvedSignature signatureOptions) {
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
      KeyStore keyStore, Path keyStorePath, ResolvedSignature signatureOptions) {
    try {
      return resolveAlias(keyStore, signatureOptions.alias().orElse(null), signatureOptions);
    } catch (GeneralSecurityException exception) {
      throw new InvalidSigningConfigurationException(
          "Failed to inspect signing aliases in " + keyStorePath, exception);
    }
  }

  /** Reads the private signing key for one resolved alias. */
  public static PrivateKey signingPrivateKey(
      KeyStore keyStore, Path keyStorePath, String alias, ResolvedSignature signatureOptions) {
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
      KeyStore keyStore, @Nullable String requestedAlias, ResolvedSignature signatureOptions)
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
