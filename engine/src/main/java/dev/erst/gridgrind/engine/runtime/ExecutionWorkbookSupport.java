package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookArtifactIo;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Workbook open, persist, and temp-file cleanup helpers for executor workflows. */
final class ExecutionWorkbookSupport {
  private final TempFileFactory tempFileFactory;

  ExecutionWorkbookSupport(TempFileFactory tempFileFactory) {
    this.tempFileFactory =
        Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
  }

  ExcelWorkbook openWorkbook(
      WorkbookPlan.WorkbookSource source,
      FormulaEnvironmentInput formulaEnvironment,
      Path workingDirectory)
      throws IOException {
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ ->
          ExcelWorkbooks.create(
              FormulaEnvironmentConverter.toExcelFormulaEnvironment(
                  formulaEnvironment, workingDirectory));
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          ExcelWorkbooks.open(
              ExecutionRequestPaths.normalizePath(existingFile.path(), workingDirectory),
              FormulaEnvironmentConverter.toExcelFormulaEnvironment(
                  formulaEnvironment, workingDirectory),
              OoxmlPackageSecurityConverter.toExcelOpenOptions(
                  existingFile.security().orElse(null)),
              tempFileFactory::createTempFile);
    };
  }

  WorkbookResultPersistence.PersistenceOutcome persistWorkbook(
      ExcelWorkbook workbook,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      Path workingDirectory)
      throws IOException {
    Objects.requireNonNull(workbook, "workbook must not be null");
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new WorkbookResultPersistence.PersistenceOutcome.NotSaved();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> {
        Path executionPath = ExecutionRequestPaths.normalizePath(saveAs.path(), workingDirectory);
        workbook
            .persistence()
            .save(
                executionPath,
                ExecutionRequestPaths.writeDisposition(saveAs),
                ExecutionRequestPaths.persistenceOptions(saveAs, workingDirectory),
                tempFileFactory::createTempFile);
        yield new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
            saveAs.path(),
            new WorkbookResultPersistence.WriteResult.Written(executionPath.toString()));
      }
      case WorkbookPlan.WorkbookPersistence.Overwrite overwrite -> {
        if (!(source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile)) {
          throw new IllegalArgumentException("OVERWRITE persistence requires an EXISTING source");
        }
        Path executionPath =
            ExecutionRequestPaths.normalizePath(existingFile.path(), workingDirectory);
        workbook
            .persistence()
            .save(
                executionPath,
                WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
                ExecutionRequestPaths.persistenceOptions(overwrite, workingDirectory),
                tempFileFactory::createTempFile);
        yield new WorkbookResultPersistence.PersistenceOutcome.Overwritten(
            existingFile.path(),
            new WorkbookResultPersistence.WriteResult.Written(executionPath.toString()));
      }
    };
  }

  WorkbookResultPersistence.PersistenceOutcome persistStreamingWorkbook(
      Path materializedPath,
      WorkbookPlan.WorkbookPersistence persistence,
      WorkbookPlan.WorkbookSource source,
      Path workingDirectory)
      throws IOException {
    Objects.requireNonNull(materializedPath, "materializedPath must not be null");
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new WorkbookResultPersistence.PersistenceOutcome.NotSaved();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> {
        Path executionPath = ExecutionRequestPaths.normalizePath(saveAs.path(), workingDirectory);
        WorkbookArtifactIo.persistMaterializedWorkbook(
            materializedPath,
            executionPath,
            ExecutionRequestPaths.sourcePackageSecurity(source),
            ExecutionRequestPaths.sourceEncryptionPassword(source),
            true,
            ExecutionRequestPaths.writeDisposition(saveAs),
            ExecutionRequestPaths.persistenceOptions(saveAs, workingDirectory));
        yield new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
            saveAs.path(),
            new WorkbookResultPersistence.WriteResult.Written(executionPath.toString()));
      }
      case WorkbookPlan.WorkbookPersistence.Overwrite overwrite -> {
        if (!(source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile)) {
          throw new IllegalArgumentException("OVERWRITE persistence requires an EXISTING source");
        }
        Path executionPath =
            ExecutionRequestPaths.normalizePath(existingFile.path(), workingDirectory);
        WorkbookArtifactIo.persistMaterializedWorkbook(
            materializedPath,
            executionPath,
            ExecutionRequestPaths.sourcePackageSecurity(source),
            ExecutionRequestPaths.sourceEncryptionPassword(source),
            true,
            WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
            ExecutionRequestPaths.persistenceOptions(overwrite, workingDirectory));
        yield new WorkbookResultPersistence.PersistenceOutcome.Overwritten(
            existingFile.path(),
            new WorkbookResultPersistence.WriteResult.Written(executionPath.toString()));
      }
    };
  }

  static void deleteIfExists(@Nullable Path path) {
    deleteIfExists(path, Files::deleteIfExists);
  }

  static void deleteIfExists(@Nullable Path path, PathDeleteOperation deleteOperation) {
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
