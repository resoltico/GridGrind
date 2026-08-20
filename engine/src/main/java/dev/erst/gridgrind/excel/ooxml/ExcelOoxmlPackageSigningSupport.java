package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.WorkbookSecurityException;
import java.io.IOException;
import java.nio.file.Path;
import javax.xml.crypto.MarshalException;
import javax.xml.crypto.dsig.XMLSignatureException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.dsig.SignatureConfig;
import org.apache.poi.poifs.crypt.dsig.SignatureInfo;

/** Loads signing material and applies validated OOXML package signatures. */
public final class ExcelOoxmlPackageSigningSupport {
  private ExcelOoxmlPackageSigningSupport() {}

  /** Removes every OOXML package signature so an unsigned or re-signed output is intentional. */
  public static void removeSignatures(Path workbookPath) throws IOException {
    ExcelOoxmlPackageSignatureRemovalSupport.removeSignatures(workbookPath);
  }

  /** Signs one materialized workbook package with the supplied PKCS#12 credentials. */
  public static void signWorkbook(Path workbookPath, ExcelOoxmlSignatureOptions signatureOptions)
      throws IOException {
    ExcelOoxmlSigningMaterialSupport.SigningMaterial signingMaterial =
        ExcelOoxmlSigningMaterialSupport.signingMaterial(signatureOptions);
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

  /** Confirms and verifies one OOXML package signature in a single testable step. */
  @FunctionalInterface
  public interface SignatureWriter {
    /** Signs and immediately verifies the saved package. */
    boolean confirmAndVerify() throws XMLSignatureException, MarshalException;
  }
}
