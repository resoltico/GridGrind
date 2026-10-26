package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.SecretReference;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/** Creates explicit host-grant documents for CLI integration tests. */
final class CliGrantFixtureSupport {
  private CliGrantFixtureSupport() {}

  static Path noPublication(List<String> operationIds) throws IOException {
    return noPublication(operationIds, List.of());
  }

  static Path noPublication(List<String> operationIds, List<Path> readablePaths)
      throws IOException {
    return write(operationIds, readablePaths, new ExecutionGrant.PublicationAuthority.None());
  }

  static Path noPublication(
      List<String> operationIds,
      List<Path> readablePaths,
      List<SecretReference> allowedSecretReferences)
      throws IOException {
    return write(
        operationIds,
        readablePaths,
        new ExecutionGrant.PublicationAuthority.None(),
        allowedSecretReferences);
  }

  static Path noPublicationWithStandardInput(List<String> operationIds) throws IOException {
    Path grantPath = Files.createTempFile("gridgrind-cli-grant-", ".json");
    grantPath.toFile().deleteOnExit();
    ExecutionGrant document =
        new ExecutionGrant(
            List.of(new ExecutionGrant.ReadAuthority.Stdin()),
            operationIds,
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            new AcceptancePolicy.MinimumOnly());
    Files.write(
        grantPath,
        dev.erst.gridgrind.json.codec.GridGrindJsonOutput.writeGrantBytes(document, false));
    return grantPath;
  }

  static Path saveAs(
      List<String> operationIds,
      List<Path> readablePaths,
      Path outputPath,
      WorkbookPlan.WorkbookPersistence.IfExists ifExists)
      throws IOException {
    return write(
        operationIds,
        readablePaths,
        new ExecutionGrant.PublicationAuthority.SaveAs(outputPath, ifExists));
  }

  static Path overwriteSource(List<String> operationIds, List<Path> readablePaths)
      throws IOException {
    return write(
        operationIds, readablePaths, new ExecutionGrant.PublicationAuthority.OverwriteSource());
  }

  private static Path write(
      List<String> operationIds,
      List<Path> readablePaths,
      ExecutionGrant.PublicationAuthority publicationAuthority)
      throws IOException {
    return write(operationIds, readablePaths, publicationAuthority, List.of());
  }

  private static Path write(
      List<String> operationIds,
      List<Path> readablePaths,
      ExecutionGrant.PublicationAuthority publicationAuthority,
      List<SecretReference> allowedSecretReferences)
      throws IOException {
    Path grantPath = Files.createTempFile("gridgrind-cli-grant-", ".json");
    grantPath.toFile().deleteOnExit();
    ExecutionGrant document =
        new ExecutionGrant(
            readablePaths.stream()
                .map(path -> new ExecutionGrant.ReadAuthority.File(path))
                .map(ExecutionGrant.ReadAuthority.class::cast)
                .toList(),
            operationIds,
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            publicationAuthority,
            allowedSecretReferences,
            new AcceptancePolicy.MinimumOnly());
    Files.write(
        grantPath,
        dev.erst.gridgrind.json.codec.GridGrindJsonOutput.writeGrantBytes(document, false));
    return grantPath;
  }
}
