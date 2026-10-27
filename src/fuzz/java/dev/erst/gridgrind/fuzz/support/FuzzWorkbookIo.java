package dev.erst.gridgrind.fuzz.support;

import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import dev.erst.gridgrind.engine.runtime.ExecutionInputBindings;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;

/** Explicit-root workbook IO helpers shared across Jazzer replay and round-trip support. */
public final class FuzzWorkbookIo {
  private static final Path MANAGED_TEMP_SEGMENT = Path.of(".gridgrind", "tmp");

  private FuzzWorkbookIo() {}

  /** Grants exactly the generated workflow's owned input, operations and publication. */
  public static ExecutionInputBindings executionBindings(
      Path workingDirectory, dev.erst.gridgrind.core.plan.WorkbookPlan request) {
    Path root = normalize(workingDirectory, "workingDirectory");
    var resources = new java.util.ArrayList<ExecutionGrant.ReadAuthority>();
    if (request.source()
        instanceof dev.erst.gridgrind.core.plan.WorkbookPlan.WorkbookSource.ExistingFile source) {
      resources.add(new ExecutionGrant.ReadAuthority.File(ownedPath(root, source.path())));
    }
    var operations = new java.util.ArrayList<String>();
    for (var step : request.steps()) {
      operations.add(
          switch (step) {
            case dev.erst.gridgrind.core.plan.MutationStep mutation ->
                dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames.mutationActionTypeName(
                    mutation.action().getClass());
            case dev.erst.gridgrind.core.plan.InspectionStep inspection ->
                dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames.inspectionQueryTypeName(
                    inspection.query().getClass());
            case dev.erst.gridgrind.core.plan.AssertionStep assertion ->
                dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames.assertionTypeName(
                    assertion.assertion().getClass());
          });
    }
    ExecutionGrant.PublicationAuthority publication =
        switch (request.persistence()) {
          case dev.erst.gridgrind.core.plan.WorkbookPlan.WorkbookPersistence.None ignored ->
              new ExecutionGrant.PublicationAuthority.None();
          case dev.erst.gridgrind.core.plan.WorkbookPlan.WorkbookPersistence.Overwrite ignored ->
              new ExecutionGrant.PublicationAuthority.OverwriteSource();
          case dev.erst.gridgrind.core.plan.WorkbookPlan.WorkbookPersistence.SaveAs save ->
              new ExecutionGrant.PublicationAuthority.SaveAs(
                  ownedPath(root, save.path()), save.ifExists());
        };
    return new ExecutionInputBindings(
        root,
        root.resolve(MANAGED_TEMP_SEGMENT),
        new ExecutionGrant(
            resources,
            operations,
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            publication,
            List.of(),
            dev.erst.gridgrind.core.grant.AcceptancePolicy.minimum()));
  }

  private static Path ownedPath(Path root, String authoredPath) {
    Path path = root.resolve(authoredPath).toAbsolutePath().normalize();
    if (!path.startsWith(root))
      throw new IllegalArgumentException("Generated workflow path escapes its owned root");
    return path;
  }

  /** Opens one workbook using a temp factory rooted next to the workbook path. */
  public static XssfWorkbook openWorkbook(Path workbookPath) throws IOException {
    return XssfWorkbooks.open(workbookPath, tempFileFactoryFor(workbookPath));
  }

  /** Saves one workbook using a temp factory rooted next to the workbook path. */
  public static void saveWorkbook(XssfWorkbook workbook, Path workbookPath) throws IOException {
    Objects.requireNonNull(workbook, "workbook must not be null");
    workbook
        .persistence()
        .save(
            workbookPath,
            WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
            tempFileFactoryFor(workbookPath));
  }

  /** Returns a temp-file factory rooted beside one anchored workbook path. */
  public static WorkbookTempFileFactory tempFileFactoryFor(Path anchoredPath) {
    Path normalizedAnchoredPath = normalize(anchoredPath, "anchoredPath");
    Path parent = normalizedAnchoredPath.getParent();
    Path workingDirectory = parent == null ? normalizedAnchoredPath : parent;
    return WorkbookTempFileFactory.rooted(workingDirectory.resolve(MANAGED_TEMP_SEGMENT));
  }

  private static Path normalize(Path path, String name) {
    return Objects.requireNonNull(path, name + " must not be null").toAbsolutePath().normalize();
  }
}
