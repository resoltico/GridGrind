package dev.erst.gridgrind.excel.ooxml;

import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.WorkbookSecurityException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Verifies the strict boundary between OOXML package parsing and signature inspection failures. */
class ExcelOoxmlPackageInspectionSupportTest {
  @Test
  void wrapsOnlyRuntimeFailuresRaisedDuringSignatureInspection() throws Exception {
    Path workbookPath = Files.createTempFile("gridgrind-signature-inspection-", ".xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        OutputStream output = Files.newOutputStream(workbookPath)) {
      workbook.createSheet("Sheet1");
      workbook.write(output);
    }

    try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ)) {
      WorkbookSecurityException failure =
          assertThrows(
              WorkbookSecurityException.class,
              () ->
                  ExcelOoxmlPackageInspectionSupport.inspectOpenedPackage(
                      pkg,
                      workbookPath,
                      ExcelOoxmlEncryptionSnapshot.none(),
                      () -> {
                        throw new IllegalStateException("signature runtime failure");
                      }));
      assertTrue(failure.getMessage().contains("inspect OOXML package signatures"));
    }
  }

  @Test
  void wrapsMalformedPackagesWhenSignatureRemovalCannotOpenThem() throws Exception {
    Path malformedWorkbookPath = Files.createTempFile("gridgrind-signature-removal-", ".xlsx");
    Files.writeString(malformedWorkbookPath, "not an OOXML package");

    WorkbookSecurityException failure =
        assertThrows(
            WorkbookSecurityException.class,
            () -> ExcelOoxmlPackageSigningSupport.removeSignatures(malformedWorkbookPath));

    assertTrue(failure.getMessage().contains("signature removal"));
  }
}
