package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySupport;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/** Workbook open, persist, and temp-file cleanup helpers for executor workflows. */
final class ExecutionWorkbookSupport {
  private final TempFileFactory tempFileFactory;

  ExecutionWorkbookSupport(TempFileFactory tempFileFactory) {
    this.tempFileFactory =
        Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
  }

  ExcelWorkbook openWorkbook(
      WorkbookPlan.WorkbookSource source, FormulaEnvironmentInput formulaEnvironment)
      throws IOException {
    return openWorkbook(source, formulaEnvironment, Path.of(""));
  }

  ExcelWorkbook openWorkbook(
      WorkbookPlan.WorkbookSource source,
      FormulaEnvironmentInput formulaEnvironment,
      Path workingDirectory)
      throws IOException {
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ ->
          ExcelWorkbook.create(
              FormulaEnvironmentConverter.toExcelFormulaEnvironment(
                  formulaEnvironment, workingDirectory));
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          ExcelWorkbook.open(
              ExecutionRequestPaths.normalizePath(existingFile.path(), workingDirectory),
              FormulaEnvironmentConverter.toExcelFormulaEnvironment(
                  formulaEnvironment, workingDirectory),
              OoxmlPackageSecurityConverter.toExcelOpenOptions(
                  existingFile.security().orElse(null)));
    };
  }

  GridGrindResponsePersistence.PersistenceOutcome persistWorkbook(
      ExcelWorkbook workbook,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence)
      throws IOException {
    return persistWorkbook(workbook, source, persistence, Path.of(""));
  }

  GridGrindResponsePersistence.PersistenceOutcome persistWorkbook(
      ExcelWorkbook workbook,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      Path workingDirectory)
      throws IOException {
    Objects.requireNonNull(workbook, "workbook must not be null");
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new GridGrindResponsePersistence.PersistenceOutcome.NotSaved();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> {
        Path executionPath = ExecutionRequestPaths.normalizePath(saveAs.path(), workingDirectory);
        workbook.save(
            executionPath,
            ExecutionRequestPaths.persistenceOptions(saveAs, workingDirectory),
            tempFileFactory::createTempFile);
        yield new GridGrindResponsePersistence.PersistenceOutcome.SavedAs(
            saveAs.path(), executionPath.toString());
      }
      case WorkbookPlan.WorkbookPersistence.OverwriteSource overwrite -> {
        if (!(source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile)) {
          throw new IllegalArgumentException("OVERWRITE persistence requires an EXISTING source");
        }
        Path executionPath =
            ExecutionRequestPaths.normalizePath(existingFile.path(), workingDirectory);
        workbook.save(
            executionPath,
            ExecutionRequestPaths.persistenceOptions(overwrite, workingDirectory),
            tempFileFactory::createTempFile);
        yield new GridGrindResponsePersistence.PersistenceOutcome.Overwritten(
            existingFile.path(), executionPath.toString());
      }
    };
  }

  GridGrindResponsePersistence.PersistenceOutcome persistStreamingWorkbook(
      Path materializedPath,
      WorkbookPlan.WorkbookPersistence persistence,
      WorkbookPlan.WorkbookSource source)
      throws IOException {
    return persistStreamingWorkbook(materializedPath, persistence, source, Path.of(""));
  }

  GridGrindResponsePersistence.PersistenceOutcome persistStreamingWorkbook(
      Path materializedPath,
      WorkbookPlan.WorkbookPersistence persistence,
      WorkbookPlan.WorkbookSource source,
      Path workingDirectory)
      throws IOException {
    Objects.requireNonNull(materializedPath, "materializedPath must not be null");
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new GridGrindResponsePersistence.PersistenceOutcome.NotSaved();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> {
        Path executionPath = ExecutionRequestPaths.normalizePath(saveAs.path(), workingDirectory);
        ExcelOoxmlPackageSecuritySupport.persistMaterializedWorkbook(
            materializedPath,
            executionPath,
            ExecutionRequestPaths.sourcePackageSecurity(source),
            ExecutionRequestPaths.sourceEncryptionPassword(source),
            true,
            ExecutionRequestPaths.persistenceOptions(saveAs, workingDirectory));
        yield new GridGrindResponsePersistence.PersistenceOutcome.SavedAs(
            saveAs.path(), executionPath.toString());
      }
      case WorkbookPlan.WorkbookPersistence.OverwriteSource overwrite -> {
        if (!(source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile)) {
          throw new IllegalArgumentException("OVERWRITE persistence requires an EXISTING source");
        }
        Path executionPath =
            ExecutionRequestPaths.normalizePath(existingFile.path(), workingDirectory);
        ExcelOoxmlPackageSecuritySupport.persistMaterializedWorkbook(
            materializedPath,
            executionPath,
            ExecutionRequestPaths.sourcePackageSecurity(source),
            ExecutionRequestPaths.sourceEncryptionPassword(source),
            true,
            ExecutionRequestPaths.persistenceOptions(overwrite, workingDirectory));
        yield new GridGrindResponsePersistence.PersistenceOutcome.Overwritten(
            existingFile.path(), executionPath.toString());
      }
    };
  }

  static void deleteIfExists(Path path) {
    deleteIfExists(path, Files::deleteIfExists);
  }

  static void deleteIfExists(Path path, PathDeleteOperation deleteOperation) {
    if (path == null) {
      return;
    }
    try {
      deleteOperation.deleteIfExists(path);
    } catch (IOException ignored) {
      // Best-effort cleanup for internal temporary files only.
    }
  }
}
