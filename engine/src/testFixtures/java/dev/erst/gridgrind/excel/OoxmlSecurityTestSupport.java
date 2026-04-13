package dev.erst.gridgrind.excel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.security.KeyPair;
import java.security.KeyPairGenerator;
import java.security.KeyStore;
import java.security.Security;
import java.security.cert.X509Certificate;
import java.time.Instant;
import java.util.Date;
import java.util.List;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.crypt.dsig.SignatureConfig;
import org.apache.poi.poifs.crypt.dsig.SignatureInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bouncycastle.jce.provider.BouncyCastleProvider;

/**
 * Deterministic signed and encrypted OOXML workbook fixtures shared across engine and protocol
 * tests.
 */
public final class OoxmlSecurityTestSupport {
  public static final String ENCRYPTION_PASSWORD = "gridgrind-encrypted-pass";
  public static final String KEYSTORE_PASSWORD = "gridgrind-keystore-pass";
  public static final String KEY_PASSWORD = "gridgrind-key-pass";
  public static final String KEY_ALIAS = "gridgrind-signing";

  private static final Instant CERT_NOT_BEFORE = Instant.parse("2026-04-13T00:00:00Z");
  private static final Instant CERT_NOT_AFTER = Instant.parse("2036-04-13T00:00:00Z");
  private static final BouncyCastleProvider BOUNCY_CASTLE_PROVIDER = new BouncyCastleProvider();

  static {
    Security.addProvider(BOUNCY_CASTLE_PROVIDER);
  }

  private OoxmlSecurityTestSupport() {}

  /** Creates one deterministic encrypted `.xlsx` fixture plus the password that unlocks it. */
  public static EncryptedWorkbook createEncryptedWorkbook(Path directory) throws IOException {
    try {
      Path fixtureDirectory = Files.createDirectories(directory);
      Path workbookPath = fixtureDirectory.resolve("encrypted.xlsx");
      byte[] workbookBytes = plainWorkbookBytes("Encrypted", "A1", "Encrypted workbook");

      EncryptionInfo encryptionInfo = new EncryptionInfo(EncryptionMode.agile);
      Encryptor encryptor = encryptionInfo.getEncryptor();
      encryptor.confirmPassword(ENCRYPTION_PASSWORD);
      try (POIFSFileSystem fileSystem = new POIFSFileSystem()) {
        try (OutputStream encryptedStream = encryptor.getDataStream(fileSystem)) {
          encryptedStream.write(workbookBytes);
        }
        try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
          fileSystem.writeFilesystem(outputStream);
        }
      }
      return new EncryptedWorkbook(workbookPath, ENCRYPTION_PASSWORD);
    } catch (GeneralSecurityException exception) {
      throw new IOException("Failed to create encrypted workbook fixture", exception);
    }
  }

  /** Creates one deterministic signed `.xlsx` fixture plus matching PKCS#12 signing material. */
  public static SignedWorkbook createSignedWorkbook(Path directory) throws IOException {
    Path fixtureDirectory = Files.createDirectories(directory);
    Path workbookPath = fixtureDirectory.resolve("signed.xlsx");
    Path pkcs12Path = fixtureDirectory.resolve("signing-material.p12");
    Files.write(workbookPath, plainWorkbookBytes("Signed", "A1", "Signed workbook"));
    try {
      SigningMaterial signingMaterial = signingMaterial();
      writePkcs12(pkcs12Path, signingMaterial);
      signWorkbook(workbookPath, signingMaterial);
    } catch (GeneralSecurityException exception) {
      throw new IOException("Failed to create signed workbook fixture", exception);
    }
    return new SignedWorkbook(workbookPath, pkcs12Path, KEYSTORE_PASSWORD, KEY_PASSWORD, KEY_ALIAS);
  }

  /** Mutates one saved workbook cell and writes the tampered copy to the requested destination. */
  public static Path tamperWorkbookCell(
      Path sourceWorkbookPath,
      Path targetWorkbookPath,
      String sheetName,
      String address,
      String value)
      throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(sourceWorkbookPath))) {
      var sheet = workbook.getSheet(sheetName);
      if (sheet == null) {
        throw new SheetNotFoundException(sheetName);
      }
      CellReference reference = new CellReference(address);
      var row = sheet.getRow(reference.getRow());
      if (row == null) {
        row = sheet.createRow(reference.getRow());
      }
      row.createCell(reference.getCol()).setCellValue(value);
      try (OutputStream outputStream = Files.newOutputStream(targetWorkbookPath)) {
        workbook.write(outputStream);
      }
      return targetWorkbookPath;
    }
  }

  /** Returns whether the OOXML package signature validates. */
  public static boolean signatureValid(Path workbookPath) throws IOException {
    try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ)) {
      SignatureInfo signatureInfo = new SignatureInfo();
      signatureInfo.setSignatureConfig(new SignatureConfig());
      signatureInfo.setOpcPackage(pkg);
      return signatureInfo.verifySignature();
    } catch (Exception exception) {
      throw new IOException("Failed to verify workbook signature: " + workbookPath, exception);
    }
  }

  /** Returns one decrypted string cell value from an encrypted workbook. */
  public static String decryptedStringCell(
      Path workbookPath, String password, String sheetName, String address) throws IOException {
    try (POIFSFileSystem fileSystem = new POIFSFileSystem(workbookPath.toFile())) {
      EncryptionInfo encryptionInfo = new EncryptionInfo(fileSystem);
      Decryptor decryptor = Decryptor.getInstance(encryptionInfo);
      if (!decryptor.verifyPassword(password)) {
        throw new IllegalArgumentException(
            "The supplied password did not decrypt the workbook fixture: " + workbookPath);
      }
      try (InputStream decryptedStream = decryptor.getDataStream(fileSystem);
          var workbook = WorkbookFactory.create(decryptedStream)) {
        return workbook
            .getSheet(sheetName)
            .getRow(new CellReference(address).getRow())
            .getCell(new CellReference(address).getCol())
            .getStringCellValue();
      }
    } catch (GeneralSecurityException exception) {
      throw new IOException("Failed to decrypt workbook fixture: " + workbookPath, exception);
    }
  }

  private static void writePkcs12(Path pkcs12Path, SigningMaterial signingMaterial)
      throws IOException, GeneralSecurityException {
    KeyStore keyStore = KeyStore.getInstance("PKCS12");
    keyStore.load(null, null);
    keyStore.setKeyEntry(
        KEY_ALIAS,
        signingMaterial.keyPair().getPrivate(),
        KEY_PASSWORD.toCharArray(),
        new java.security.cert.Certificate[] {signingMaterial.certificate()});
    try (OutputStream outputStream = Files.newOutputStream(pkcs12Path)) {
      keyStore.store(outputStream, KEYSTORE_PASSWORD.toCharArray());
    }
  }

  private static void signWorkbook(Path workbookPath, SigningMaterial signingMaterial)
      throws IOException {
    boolean valid;
    try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ_WRITE)) {
      SignatureConfig signatureConfig = new SignatureConfig();
      signatureConfig.setExecutionTime(Date.from(CERT_NOT_BEFORE));
      signatureConfig.setKey(signingMaterial.keyPair().getPrivate());
      signatureConfig.setSigningCertificateChain(List.of(signingMaterial.certificate()));

      SignatureInfo signatureInfo = new SignatureInfo();
      signatureInfo.setSignatureConfig(signatureConfig);
      signatureInfo.setOpcPackage(pkg);
      signatureInfo.confirmSignature();
      valid = signatureInfo.verifySignature();
    } catch (javax.xml.crypto.MarshalException
        | javax.xml.crypto.dsig.XMLSignatureException
        | org.apache.poi.openxml4j.exceptions.InvalidFormatException exception) {
      throw new IOException("Failed to sign workbook fixture: " + workbookPath, exception);
    } catch (RuntimeException exception) {
      throw new IOException("Failed to sign workbook fixture: " + workbookPath, exception);
    }
    if (!valid) {
      throw new IOException("Signed workbook fixture must validate immediately");
    }
  }

  private static byte[] plainWorkbookBytes(String sheetName, String address, String value)
      throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
      var sheet = workbook.createSheet(sheetName);
      CellReference reference = new CellReference(address);
      var row = sheet.createRow(reference.getRow());
      row.createCell(reference.getCol()).setCellValue(value);
      workbook.write(outputStream);
      return outputStream.toByteArray();
    }
  }

  private static SigningMaterial signingMaterial() throws GeneralSecurityException {
    KeyPairGenerator keyPairGenerator = KeyPairGenerator.getInstance("RSA");
    keyPairGenerator.initialize(2048);
    KeyPair keyPair = keyPairGenerator.generateKeyPair();
    var subject = new org.bouncycastle.asn1.x500.X500Name("CN=GridGrind Signing Test");
    var certificateBuilder =
        new org.bouncycastle.cert.jcajce.JcaX509v3CertificateBuilder(
            subject,
            BigInteger.valueOf(20260413L),
            Date.from(CERT_NOT_BEFORE),
            Date.from(CERT_NOT_AFTER),
            subject,
            keyPair.getPublic());
    var signer = contentSigner(keyPair);
    var holder = certificateBuilder.build(signer);
    X509Certificate certificate = certificate(holder);
    certificate.checkValidity(Date.from(CERT_NOT_BEFORE.plusSeconds(1)));
    certificate.verify(keyPair.getPublic());
    return new SigningMaterial(keyPair, certificate);
  }

  private static org.bouncycastle.operator.ContentSigner contentSigner(KeyPair keyPair)
      throws GeneralSecurityException {
    try {
      return new org.bouncycastle.operator.jcajce.JcaContentSignerBuilder("SHA256withRSA")
          .setProvider(BOUNCY_CASTLE_PROVIDER)
          .build(keyPair.getPrivate());
    } catch (org.bouncycastle.operator.OperatorCreationException exception) {
      throw new GeneralSecurityException(
          "Failed to construct the PKCS#12 fixture signer", exception);
    }
  }

  private static X509Certificate certificate(org.bouncycastle.cert.X509CertificateHolder holder)
      throws GeneralSecurityException {
    return new org.bouncycastle.cert.jcajce.JcaX509CertificateConverter()
        .setProvider(BOUNCY_CASTLE_PROVIDER)
        .getCertificate(holder);
  }

  /** One encrypted workbook fixture and its unlock password. */
  public record EncryptedWorkbook(Path workbookPath, String password) {
    public EncryptedWorkbook {
      workbookPath = workbookPath.toAbsolutePath().normalize();
    }
  }

  /** One signed workbook fixture plus matching PKCS#12 signing material. */
  public record SignedWorkbook(
      Path workbookPath,
      Path pkcs12Path,
      String keystorePassword,
      String keyPassword,
      String alias) {
    public SignedWorkbook {
      workbookPath = workbookPath.toAbsolutePath().normalize();
      pkcs12Path = pkcs12Path.toAbsolutePath().normalize();
    }
  }

  private record SigningMaterial(KeyPair keyPair, X509Certificate certificate) {}
}
