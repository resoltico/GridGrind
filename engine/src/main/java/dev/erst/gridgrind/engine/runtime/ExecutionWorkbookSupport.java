package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.excel.ExcelTempFileWriteTargetSupport;
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
      @Nullable FormulaEnvironmentInput formulaEnvironment,
      ExecutionInputBindings bindings)
      throws IOException {
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ ->
          ExcelWorkbooks.create(
              FormulaEnvironmentConverter.toExcelFormulaEnvironment(formulaEnvironment, bindings));
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile -> {
        Path sourcePath =
            ExecutionRequestPaths.normalizePath(existingFile.path(), bindings.workingDirectory());
        Path materializedSource;
        try {
          materializedSource =
              bindings
                  .requestPathAccess()
                  .materializeRead(
                      existingFile.path(), "source", "gridgrind-source-workbook-", ".xlsx");
        } catch (java.nio.file.NoSuchFileException exception) {
          throw new dev.erst.gridgrind.excel.WorkbookNotFoundException(sourcePath, exception);
        }
        yield ExcelWorkbooks.open(
            materializedSource,
            FormulaEnvironmentConverter.toExcelFormulaEnvironment(formulaEnvironment, bindings),
            OoxmlPackageSecurityConverter.toExcelOpenOptions(existingFile.security().orElse(null)),
            tempFileFactory::createTempFile);
      }
    };
  }

  WorkbookResultPersistence.PersistenceOutcome persistWorkbook(
      ExcelWorkbook workbook,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionInputBindings bindings)
      throws IOException {
    Objects.requireNonNull(workbook, "workbook must not be null");
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new WorkbookResultPersistence.PersistenceOutcome.NotSaved();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> {
        Path executionPath = bindings.requestPathAccess().outputPath();
        persistToBoundOutput(
            createStagingTarget(),
            stagedPath ->
                workbook
                    .persistence()
                    .save(
                        stagedPath,
                        WorkbookArtifactWriteDisposition.CREATE_NEW,
                        ExecutionRequestPaths.persistenceOptions(saveAs, bindings),
                        tempFileFactory::createTempFile),
            bindings,
            ExecutionRequestPaths.writeDisposition(saveAs));
        yield new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
            saveAs.path(),
            new WorkbookResultPersistence.WriteResult.Written(executionPath.toString()));
      }
      case WorkbookPlan.WorkbookPersistence.Overwrite overwrite -> {
        if (!(source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile)) {
          throw new IllegalArgumentException("OVERWRITE persistence requires an EXISTING source");
        }
        Path executionPath = bindings.requestPathAccess().outputPath();
        persistToBoundOutput(
            createStagingTarget(),
            stagedPath ->
                workbook
                    .persistence()
                    .save(
                        stagedPath,
                        WorkbookArtifactWriteDisposition.CREATE_NEW,
                        ExecutionRequestPaths.persistenceOptions(overwrite, bindings),
                        tempFileFactory::createTempFile),
            bindings,
            WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
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
      ExecutionInputBindings bindings)
      throws IOException {
    Objects.requireNonNull(materializedPath, "materializedPath must not be null");
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new WorkbookResultPersistence.PersistenceOutcome.NotSaved();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> {
        Path executionPath = bindings.requestPathAccess().outputPath();
        persistToBoundOutput(
            createStagingTarget(),
            stagedPath ->
                WorkbookArtifactIo.persistMaterializedWorkbook(
                    materializedPath,
                    stagedPath,
                    ExecutionRequestPaths.sourcePackageSecurity(source),
                    ExecutionRequestPaths.sourceEncryptionPassword(source),
                    WorkbookArtifactWriteDisposition.CREATE_NEW,
                    ExecutionRequestPaths.persistenceOptions(saveAs, bindings)),
            bindings,
            ExecutionRequestPaths.writeDisposition(saveAs));
        yield new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
            saveAs.path(),
            new WorkbookResultPersistence.WriteResult.Written(executionPath.toString()));
      }
      case WorkbookPlan.WorkbookPersistence.Overwrite overwrite -> {
        if (!(source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile)) {
          throw new IllegalArgumentException("OVERWRITE persistence requires an EXISTING source");
        }
        Path executionPath = bindings.requestPathAccess().outputPath();
        persistToBoundOutput(
            createStagingTarget(),
            stagedPath ->
                WorkbookArtifactIo.persistMaterializedWorkbook(
                    materializedPath,
                    stagedPath,
                    ExecutionRequestPaths.sourcePackageSecurity(source),
                    ExecutionRequestPaths.sourceEncryptionPassword(source),
                    WorkbookArtifactWriteDisposition.CREATE_NEW,
                    ExecutionRequestPaths.persistenceOptions(overwrite, bindings)),
            bindings,
            WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
        yield new WorkbookResultPersistence.PersistenceOutcome.Overwritten(
            existingFile.path(),
            new WorkbookResultPersistence.WriteResult.Written(executionPath.toString()));
      }
    };
  }

  private Path createStagingTarget() throws IOException {
    return ExcelTempFileWriteTargetSupport.prepareCreateNewTarget(
        tempFileFactory.createTempFile("gridgrind-persistence-stage-", ".xlsx"));
  }

  private static void persistToBoundOutput(
      Path stagedPath,
      StagedWorkbookWriter writer,
      ExecutionInputBindings bindings,
      WorkbookArtifactWriteDisposition disposition)
      throws IOException {
    try {
      writer.write(stagedPath);
      bindings.requestPathAccess().commitOutput(stagedPath, disposition);
    } finally {
      deleteIfExists(stagedPath);
    }
  }

  /** Writes one complete workbook package to an executor-private staging path. */
  @FunctionalInterface
  private interface StagedWorkbookWriter {
    /** Writes the staged workbook package. */
    void write(Path stagedPath) throws IOException;
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
