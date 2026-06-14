package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.WorkbookSecurityException;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import java.io.IOException;
import java.nio.file.Path;
import java.security.cert.X509Certificate;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Optional;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.dsig.SignatureConfig;
import org.apache.poi.poifs.crypt.dsig.SignatureInfo;
import org.apache.poi.poifs.crypt.dsig.SignaturePart;
import org.jspecify.annotations.Nullable;

/** Inspects OOXML package signatures and projects them into GridGrind snapshot types. */
public final class ExcelOoxmlPackageInspectionSupport {
  private ExcelOoxmlPackageInspectionSupport() {}

  /** Reads one workbook package's signature state alongside its already-known encryption facts. */
  public static ExcelOoxmlPackageSecuritySnapshot inspectPackageSecurity(
      Path workbookPath, ExcelOoxmlEncryptionSnapshot encryption) throws IOException {
    try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ)) {
      SignatureInfo signatureInfo = ExcelOoxmlDsigProviderSupport.newSignatureInfo();
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

  /** Converts one POI signature part into the public GridGrind signature snapshot. */
  public static ExcelOoxmlSignatureSnapshot signatureSnapshot(SignaturePart signaturePart) {
    X509Certificate signerCertificate = signaturePart.getSigner();
    return new ExcelOoxmlSignatureSnapshot(
        signaturePart.getPackagePart().getPartName().getName(),
        signerIdentity(signerCertificate),
        signaturePart.validate()
            ? ExcelOoxmlSignatureState.VALID
            : ExcelOoxmlSignatureState.INVALID);
  }

  /** Returns the factual signer identity when one certificate is present. */
  public static Optional<ExcelOoxmlSignatureSnapshot.SignerIdentity> signerIdentity(
      @Nullable X509Certificate signer) {
    return signer == null
        ? Optional.empty()
        : Optional.of(
            new ExcelOoxmlSignatureSnapshot.SignerIdentity(
                signer.getSubjectX500Principal().getName(),
                signer.getIssuerX500Principal().getName(),
                signer.getSerialNumber().toString(16).toUpperCase(Locale.ROOT)));
  }
}
