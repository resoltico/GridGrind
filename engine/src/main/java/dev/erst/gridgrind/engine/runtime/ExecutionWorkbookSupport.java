package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.excel.ExcelTempFileWriteTargetSupport;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookArtifactIo;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlOpenOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceEncryption;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceOptions;
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
        try {
          yield ExcelWorkbooks.open(
              materializedSource,
              FormulaEnvironmentConverter.toExcelFormulaEnvironment(formulaEnvironment, bindings),
              OoxmlPackageSecurityConverter.toExcelOpenOptions(
                  existingFile.security().orElse(null), bindings),
              tempFileFactory::createTempFile);
        } catch (dev.erst.gridgrind.excel.WorkbookNotOpenableException exception) {
          throw new dev.erst.gridgrind.excel.WorkbookNotOpenableException(sourcePath, exception);
        }
      }
    };
  }

  WorkbookResultPersistence.PersistenceOutcome persistWorkbook(
      ExcelWorkbook workbook,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionInputBindings bindings,
      StagedArtifactAcceptance stagedArtifactAcceptance)
      throws IOException {
    Objects.requireNonNull(workbook, "workbook must not be null");
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new WorkbookResultPersistence.PersistenceOutcome.NotSaved();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> {
        WorkbookResultPersistence.PublicationOutcome.Published publication =
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
                source,
                saveAs,
                ExecutionRequestPaths.writeDisposition(saveAs),
                stagedArtifactAcceptance);
        yield new WorkbookResultPersistence.PersistenceOutcome.SavedAs(saveAs.path(), publication);
      }
      case WorkbookPlan.WorkbookPersistence.Overwrite overwrite -> {
        if (!(source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile)) {
          throw new IllegalArgumentException("OVERWRITE persistence requires an EXISTING source");
        }
        WorkbookResultPersistence.PublicationOutcome.Published publication =
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
                source,
                overwrite,
                WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
                stagedArtifactAcceptance);
        yield new WorkbookResultPersistence.PersistenceOutcome.Overwritten(
            existingFile.path(), publication);
      }
    };
  }

  WorkbookResultPersistence.PersistenceOutcome persistStreamingWorkbook(
      Path materializedPath,
      WorkbookPlan.WorkbookPersistence persistence,
      WorkbookPlan.WorkbookSource source,
      ExecutionInputBindings bindings,
      StagedArtifactAcceptance stagedArtifactAcceptance)
      throws IOException {
    Objects.requireNonNull(materializedPath, "materializedPath must not be null");
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new WorkbookResultPersistence.PersistenceOutcome.NotSaved();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> {
        WorkbookResultPersistence.PublicationOutcome.Published publication =
            persistToBoundOutput(
                createStagingTarget(),
                stagedPath ->
                    WorkbookArtifactIo.persistMaterializedWorkbook(
                        materializedPath,
                        stagedPath,
                        ExecutionRequestPaths.sourcePackageSecurity(source),
                        ExecutionRequestPaths.sourceEncryptionPassword(source, bindings),
                        WorkbookArtifactWriteDisposition.CREATE_NEW,
                        ExecutionRequestPaths.persistenceOptions(saveAs, bindings)),
                bindings,
                source,
                saveAs,
                ExecutionRequestPaths.writeDisposition(saveAs),
                stagedArtifactAcceptance);
        yield new WorkbookResultPersistence.PersistenceOutcome.SavedAs(saveAs.path(), publication);
      }
      case WorkbookPlan.WorkbookPersistence.Overwrite overwrite -> {
        if (!(source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile)) {
          throw new IllegalArgumentException("OVERWRITE persistence requires an EXISTING source");
        }
        WorkbookResultPersistence.PublicationOutcome.Published publication =
            persistToBoundOutput(
                createStagingTarget(),
                stagedPath ->
                    WorkbookArtifactIo.persistMaterializedWorkbook(
                        materializedPath,
                        stagedPath,
                        ExecutionRequestPaths.sourcePackageSecurity(source),
                        ExecutionRequestPaths.sourceEncryptionPassword(source, bindings),
                        WorkbookArtifactWriteDisposition.CREATE_NEW,
                        ExecutionRequestPaths.persistenceOptions(overwrite, bindings)),
                bindings,
                source,
                overwrite,
                WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
                stagedArtifactAcceptance);
        yield new WorkbookResultPersistence.PersistenceOutcome.Overwritten(
            existingFile.path(), publication);
      }
    };
  }

  private Path createStagingTarget() throws IOException {
    return ExcelTempFileWriteTargetSupport.prepareCreateNewTarget(
        tempFileFactory.createTempFile("gridgrind-persistence-stage-", ".xlsx"));
  }

  private WorkbookResultPersistence.PublicationOutcome.Published persistToBoundOutput(
      Path stagedPath,
      StagedWorkbookWriter writer,
      ExecutionInputBindings bindings,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      WorkbookArtifactWriteDisposition disposition,
      StagedArtifactAcceptance stagedArtifactAcceptance)
      throws IOException {
    try {
      writer.write(stagedPath);
      verifyStagedArtifact(stagedPath, source, persistence, bindings, stagedArtifactAcceptance);
      return bindings
          .requestPathAccess()
          .publishOutput(
              stagedPath,
              disposition,
              new WorkbookResultPersistence.PublicationOutcome.StagedArtifactVerification());
    } finally {
      deleteIfExists(stagedPath);
    }
  }

  private void verifyStagedArtifact(
      Path stagedPath,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionInputBindings bindings,
      StagedArtifactAcceptance stagedArtifactAcceptance)
      throws IOException {
    ExcelOoxmlPersistenceOptions persistenceOptions =
        ExecutionRequestPaths.persistenceOptions(persistence, bindings);
    try (WorkbookArtifactIo.MaterializedWorkbook materialized =
        WorkbookArtifactIo.materializeWorkbook(
            stagedPath,
            stagedArtifactOpenOptions(persistenceOptions, source, bindings),
            tempFileFactory::createTempFile)) {
      requireRegularFile(materialized.workbookPath());
      try (ExcelWorkbook reopened =
          ExcelWorkbooks.open(materialized.workbookPath(), tempFileFactory::createTempFile)) {
        reopened.sheets().sheetCount();
        Objects.requireNonNull(
                stagedArtifactAcceptance, "stagedArtifactAcceptance must not be null")
            .verify(materialized.workbookPath());
      }
    }
  }

  /** Requires a materialized staged artifact to be a regular file before workbook reopening. */
  static void requireRegularFile(Path materializedWorkbook) throws IOException {
    if (!Files.isRegularFile(materializedWorkbook)) {
      throw new IOException("staged workbook materialization did not produce a regular file");
    }
  }

  /** Derives the only valid reopen options for an already serialized staged artifact. */
  static ExcelOoxmlOpenOptions stagedArtifactOpenOptions(
      ExcelOoxmlPersistenceOptions persistenceOptions,
      WorkbookPlan.WorkbookSource source,
      ExecutionInputBindings bindings) {
    return switch (persistenceOptions.encryption()) {
      case ExcelOoxmlPersistenceEncryption.Plaintext _ -> new ExcelOoxmlOpenOptions.Unencrypted();
      case ExcelOoxmlPersistenceEncryption.Encrypt encrypt ->
          new ExcelOoxmlOpenOptions.Encrypted(encrypt.options().password());
      case ExcelOoxmlPersistenceEncryption.PreserveSource _ ->
          new ExcelOoxmlOpenOptions.Encrypted(
              ExecutionRequestPaths.sourceEncryptionPassword(source, bindings)
                  .orElseThrow(
                      () ->
                          new IllegalStateException(
                              "preserved source encryption requires a verified source password")));
    };
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
