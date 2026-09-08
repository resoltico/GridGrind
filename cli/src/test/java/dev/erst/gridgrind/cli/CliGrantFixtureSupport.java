package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.SecretReference;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
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
    return write(operationIds, readablePaths, new CliGrantPublicationAuthority.None());
  }

  static Path noPublication(
      List<String> operationIds,
      List<Path> readablePaths,
      List<SecretReference> allowedSecretReferences)
      throws IOException {
    return write(
        operationIds,
        readablePaths,
        new CliGrantPublicationAuthority.None(),
        allowedSecretReferences);
  }

  static Path noPublicationWithStandardInput(List<String> operationIds) throws IOException {
    Path grantPath = Files.createTempFile("gridgrind-cli-grant-", ".json");
    CliExecutionGrantDocument document =
        new CliExecutionGrantDocument(
            List.of(new CliExecutionGrantDocument.ReadableResource.StandardInput()),
            operationIds,
            new CliGrantTargetAuthority.WorkbookWide(),
            new CliGrantPublicationAuthority.None(),
            List.of(),
            new CliGrantAcceptancePolicy.MinimumOnly());
    Files.write(grantPath, GridGrindCliJson.writeBytes(document));
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
        new CliGrantPublicationAuthority.SaveAs(outputPath.toString(), ifExists));
  }

  static Path overwriteSource(List<String> operationIds, List<Path> readablePaths)
      throws IOException {
    return write(operationIds, readablePaths, new CliGrantPublicationAuthority.OverwriteSource());
  }

  private static Path write(
      List<String> operationIds,
      List<Path> readablePaths,
      CliGrantPublicationAuthority publicationAuthority)
      throws IOException {
    return write(operationIds, readablePaths, publicationAuthority, List.of());
  }

  private static Path write(
      List<String> operationIds,
      List<Path> readablePaths,
      CliGrantPublicationAuthority publicationAuthority,
      List<SecretReference> allowedSecretReferences)
      throws IOException {
    Path grantPath = Files.createTempFile("gridgrind-cli-grant-", ".json");
    CliExecutionGrantDocument document =
        new CliExecutionGrantDocument(
            readablePaths.stream()
                .map(path -> new CliExecutionGrantDocument.ReadableResource.File(path.toString()))
                .map(CliExecutionGrantDocument.ReadableResource.class::cast)
                .toList(),
            operationIds,
            new CliGrantTargetAuthority.WorkbookWide(),
            publicationAuthority,
            allowedSecretReferences,
            new CliGrantAcceptancePolicy.MinimumOnly());
    Files.write(grantPath, GridGrindCliJson.writeBytes(document));
    return grantPath;
  }
}
