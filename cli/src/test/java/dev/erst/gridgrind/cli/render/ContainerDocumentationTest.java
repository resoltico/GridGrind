package dev.erst.gridgrind.cli.render;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.catalog.RepositoryRootTestSupport;
import java.io.IOException;
import java.nio.file.Files;
import org.junit.jupiter.api.Test;

/** Checks container-host guidance against its delivered command rendering. */
class ContainerDocumentationTest {
  @Test
  void dockerDocsFollowThePublishedMountedWorkdirContract() throws IOException {
    String operations = readDoc("docs/OPERATIONS.md");
    String quickStart = readDoc("docs/QUICK_START.md");
    String readme = readDoc("README.md");
    String normalizedOperations = normalizeMarkdownText(operations).replace("`", "");
    String normalizedQuickStart = normalizeMarkdownText(quickStart);
    String normalizedReadme = normalizeMarkdownText(readme);
    String expectedMountSummary = ContainerHelp.dockerMountedWorkdirSummary();
    String expectedUserArgument = ContainerHelp.dockerMountedWorkdirUserArgument();
    String expectedExecutionCommand =
        ContainerHelp.dockerMountedWorkdirExecutionCommand("ghcr.io/resoltico/gridgrind:latest");
    String expectedVolumeArgument = ContainerHelp.dockerMountedWorkdirVolumeArgument();
    String expectedReadmeExecutionCommand =
        "docker run --pull=always --rm -i "
            + expectedUserArgument
            + " "
            + expectedVolumeArgument
            + " ghcr.io/resoltico/gridgrind:latest run --request request.json --grant grant.json --response"
            + " response.json";
    String expectedQuickStartReleaseExecutionCommand =
        "docker run --pull=always --rm -i "
            + expectedUserArgument
            + " "
            + expectedVolumeArgument
            + " ghcr.io/resoltico/gridgrind:latest run --request budget-request.json --grant grant.json --response"
            + " response.json";
    String expectedQuickStartLocalExecutionCommand =
        "docker run --rm -i "
            + expectedUserArgument
            + " "
            + expectedVolumeArgument
            + " gridgrind-local run --request budget-request.json --grant grant.json --response response.json";

    assertAll(
        () ->
            assertTrue(
                normalizedOperations.contains(expectedMountSummary),
                "operations doc must use the canonical mounted-workdir Docker wording"),
        () ->
            assertTrue(
                operations.contains(expectedUserArgument),
                "operations doc must publish the canonical bind-mount --user argument"),
        () ->
            assertTrue(
                operations.contains(expectedExecutionCommand),
                "operations doc must publish the canonical mounted-directory Docker command"),
        () ->
            assertTrue(
                operations.contains(expectedVolumeArgument),
                "operations doc must mount the host working directory through the canonical /work volume argument"),
        () ->
            assertFalse(
                operations.contains("-w /workdir"),
                "operations doc must not revive the removed Docker -w override"),
        () ->
            assertTrue(
                normalizedQuickStart.contains(expectedQuickStartReleaseExecutionCommand),
                "quick start must publish the exact released-image mounted-directory Docker"
                    + " command"),
        () ->
            assertTrue(
                normalizedQuickStart.contains(expectedQuickStartLocalExecutionCommand),
                "quick start must publish the exact local-image mounted-directory Docker"
                    + " command"),
        () ->
            assertFalse(
                quickStart.contains("-w /workdir"),
                "quick start must not teach the removed Docker -w override"),
        () ->
            assertTrue(
                normalizedReadme.contains(expectedReadmeExecutionCommand),
                "readme must keep the published Docker bind-mount command in sync"),
        () ->
            assertFalse(
                readme.contains("-w /workdir"),
                "readme must not revive the removed Docker -w override"));
  }

  private static String readDoc(String relativePath) throws IOException {
    return Files.readString(RepositoryRootTestSupport.repositoryRoot().resolve(relativePath));
  }

  private static String normalizeMarkdownText(String markdown) {
    return markdown.replace("\\", "").replaceAll("\\s+", " ").trim();
  }
}
