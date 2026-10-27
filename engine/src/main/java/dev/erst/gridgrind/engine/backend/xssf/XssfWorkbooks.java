package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.model.security.ResolvedOpenSecurity;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import dev.erst.gridgrind.engine.security.SecurityAdapter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Factory and open surface for materializing workbook wrappers. */
public final class XssfWorkbooks {
  private XssfWorkbooks() {}

  /** Creates an empty XLSX workbook. */
  public static XssfWorkbook create() {
    return new XssfWorkbook(new XSSFWorkbook());
  }

  /** Wraps one already-materialized POI workbook inside the GridGrind workbook boundary. */
  public static XssfWorkbook wrap(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    return new XssfWorkbook(workbook);
  }

  /** Creates an empty XLSX workbook with the supplied formula-evaluation environment. */
  public static XssfWorkbook create(ResolvedFormulaEnvironment formulaEnvironment)
      throws IOException {
    return new XssfWorkbook(new XSSFWorkbook(), formulaEnvironment);
  }

  /** Opens an existing workbook file from disk with explicit temp-file ownership. */
  public static XssfWorkbook open(Path workbookPath, WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    return open(workbookPath, new ResolvedOpenSecurity.Unencrypted(), tempFileFactory);
  }

  /**
   * Opens an existing workbook with explicit package-open settings and a custom temp-file factory.
   */
  public static XssfWorkbook open(
      Path workbookPath, ResolvedOpenSecurity openOptions, WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    return SecurityAdapter.openWorkbook(workbookPath, openOptions, tempFileFactory);
  }

  /** Opens an existing workbook file from disk with formula evaluation and temp ownership. */
  public static XssfWorkbook open(
      Path workbookPath,
      ResolvedFormulaEnvironment formulaEnvironment,
      WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    return open(
        workbookPath, formulaEnvironment, new ResolvedOpenSecurity.Unencrypted(), tempFileFactory);
  }

  /**
   * Opens an existing workbook with formula and package-open settings plus a custom temp-file
   * factory.
   */
  public static XssfWorkbook open(
      Path workbookPath,
      ResolvedFormulaEnvironment formulaEnvironment,
      ResolvedOpenSecurity openOptions,
      WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    return SecurityAdapter.openWorkbook(
        workbookPath, formulaEnvironment, openOptions, tempFileFactory);
  }

  /** Maximum size of a single decompressed ZIP entry in an xlsx package (LIM-026). */
  static final long MAX_ZIP_ENTRY_SIZE = 100L * 1024 * 1024; // 100 MiB

  static {
    ZipSecureFile.setMinInflateRatio(0.01d); // LIM-026: 1:100 max decompression ratio
    ZipSecureFile.setMaxEntrySize(MAX_ZIP_ENTRY_SIZE); // LIM-026
  }

  /** Opens a plain materialized workbook with retained source-security metadata. */
  public static XssfWorkbook openMaterializedWorkbook(
      Path workbookPath,
      Optional<Path> sourcePath,
      OoxmlPackageSecurityReport loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword)
      throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(loadedPackageSecurity, "loadedPackageSecurity must not be null");

    Path absolutePath = workbookPath.toAbsolutePath();
    if (!Files.exists(absolutePath)) {
      throw new WorkbookNotFoundException(absolutePath);
    }

    try (InputStream inputStream = Files.newInputStream(absolutePath)) {
      try {
        return openMaterializedWorkbook(
            new XSSFWorkbook(inputStream),
            sourcePath,
            loadedPackageSecurity,
            sourceEncryptionPassword);
      } catch (NotOfficeXmlFileException exception) {
        throw new IllegalArgumentException("Only .xlsx workbooks are supported", exception);
      }
    }
  }

  /** Opens a plain materialized workbook with retained source-security metadata. */
  public static XssfWorkbook openMaterializedWorkbook(
      Path workbookPath,
      ResolvedFormulaEnvironment formulaEnvironment,
      Optional<Path> sourcePath,
      OoxmlPackageSecurityReport loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword)
      throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(loadedPackageSecurity, "loadedPackageSecurity must not be null");

    Path absolutePath = workbookPath.toAbsolutePath();
    if (!Files.exists(absolutePath)) {
      throw new WorkbookNotFoundException(absolutePath);
    }

    try (InputStream inputStream = Files.newInputStream(absolutePath)) {
      try {
        return openMaterializedWorkbook(
            new XSSFWorkbook(inputStream),
            formulaEnvironment,
            sourcePath,
            loadedPackageSecurity,
            sourceEncryptionPassword);
      } catch (NotOfficeXmlFileException exception) {
        throw new IllegalArgumentException("Only .xlsx workbooks are supported", exception);
      }
    }
  }

  static XssfWorkbook openMaterializedWorkbook(
      XSSFWorkbook xssfWorkbook,
      Optional<Path> sourcePath,
      OoxmlPackageSecurityReport loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword)
      throws IOException {
    try {
      return new XssfWorkbook(
          xssfWorkbook,
          FormulaRuntime.poi(xssfWorkbook.getCreationHelper().createFormulaEvaluator()),
          sourcePath,
          loadedPackageSecurity,
          sourceEncryptionPassword);
    } catch (RuntimeException exception) {
      closeWorkbookAfterOpenFailure(xssfWorkbook, exception);
      throw exception;
    }
  }

  static XssfWorkbook openMaterializedWorkbook(
      XSSFWorkbook xssfWorkbook,
      ResolvedFormulaEnvironment formulaEnvironment,
      Optional<Path> sourcePath,
      OoxmlPackageSecurityReport loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword)
      throws IOException {
    try {
      return new XssfWorkbook(
          xssfWorkbook,
          FormulaRuntime.poi(xssfWorkbook, formulaEnvironment),
          sourcePath,
          loadedPackageSecurity,
          sourceEncryptionPassword);
    } catch (IOException | RuntimeException exception) {
      closeWorkbookAfterOpenFailure(xssfWorkbook, exception);
      throw exception;
    }
  }

  static void closeWorkbookAfterOpenFailure(XSSFWorkbook workbook, Exception exception)
      throws IOException {
    try {
      workbook.close();
    } catch (IOException closeException) {
      exception.addSuppressed(closeException);
    }
  }
}
