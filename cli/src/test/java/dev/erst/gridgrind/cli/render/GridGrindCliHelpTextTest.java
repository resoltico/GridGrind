package dev.erst.gridgrind.cli.render;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.command.CliCommand;
import dev.erst.gridgrind.cli.command.GridGrindCli;
import dev.erst.gridgrind.cli.command.GridGrindCliTestSupport;
import dev.erst.gridgrind.core.discovery.ShippedExampleEntry;
import dev.erst.gridgrind.core.discovery.ShippedExamples;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Properties;
import org.junit.jupiter.api.Test;

/** Help, version, and documentation integration tests for GridGrindCli. */
class GridGrindCliHelpTextTest extends GridGrindCliTestSupport {
  @Test
  void bundledProductPropertiesMatchTheCurrentGradleProperties() throws IOException {
    Properties properties = new Properties();
    try (InputStream inputStream =
        Files.newInputStream(locateRepoRoot().resolve("gradle.properties"))) {
      properties.load(inputStream);
    }

    assertEquals(properties.getProperty("version"), ProductInfo.version());
    assertEquals(properties.getProperty("gridgrindDescription"), ProductInfo.description());
  }

  @Test
  void versionFromReturnsBundledVersionWhenImplementationVersionIsAbsent() throws IOException {
    assertEquals(ProductInfo.version(), ProductInfo.versionFrom(null));
  }

  @Test
  void versionFromReturnsUnknownWhenBundledVersionIsUnavailable() {
    assertEquals("unknown", ProductInfo.versionFrom(null, Object.class));
  }

  @Test
  void versionFromIgnoresBlankImplementationVersionAndBlankBundledVersion() {
    assertEquals(
        "unknown",
        ProductInfo.versionFrom(
            "   ", new ByteArrayInputStream("version=\n".getBytes(StandardCharsets.UTF_8))));
  }

  @Test
  void versionFromReturnsVersionWhenImplementationVersionIsPresent() {
    assertEquals("0.4.1", ProductInfo.versionFrom("0.4.1"));
  }

  @Test
  void versionFlagPrintsVersionLineToStdoutAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"version"}, new ByteArrayInputStream(new byte[0]), stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8);
    assertTrue(output.startsWith("GridGrind " + ProductInfo.version() + "\n"));
    assertTrue(output.endsWith("\n"));
    assertTrue(output.lines().count() >= 2);
  }

  @Test
  void licenseFlagPrintsLicenseTextToStdoutAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"license"}, new ByteArrayInputStream(new byte[0]), stdout);

    assertEquals(0, exitCode);
    assertFalse(stdout.toString(StandardCharsets.UTF_8).isBlank());
  }

  @Test
  void descriptionFromReturnsFallbackWhenPropertiesAreUnavailableOrBlank() throws IOException {
    assertEquals("GridGrind", ProductInfo.descriptionFrom(Object.class));
    assertEquals(
        "GridGrind",
        ProductInfo.descriptionFrom(
            new ByteArrayInputStream("description=\n".getBytes(StandardCharsets.UTF_8))));
  }

  @Test
  void licenseTextContainsMplLicenseWhenResourcePresent() {
    String mpl = "Mozilla Public License Version 2.0\n\nCopyright (c) 2026 Ervins Strauhmanis\n";
    InputStream own = new ByteArrayInputStream(mpl.getBytes(StandardCharsets.UTF_8));

    String result = ProductInfo.licenseText(own, null, null, null, null);

    assertTrue(result.contains("Mozilla Public License Version 2.0"));
    assertTrue(result.contains("Ervins Strauhmanis"));
    assertFalse(result.contains("Third-party notices and licenses:"));
  }

  @Test
  void packagedLicenseTextCoversResolvedRuntimeTerms() {
    String result = ProductInfo.licenseText(GridGrindCli.class);

    assertTrue(result.contains("Mozilla Public License Version 2.0"));
    assertTrue(result.contains("https://github.com/resoltico/GridGrind/releases"));
    assertTrue(result.contains("Copyright (c) 2026 Ervins Strauhmanis"));
    assertTrue(result.contains("Jakarta Activation API"));
    assertTrue(result.contains("Jakarta XML Binding API"));
    assertTrue(result.contains("BSD 3-Clause"));
    assertTrue(result.contains("Eclipse Distribution License v1.0"));
    assertTrue(result.contains("FastDoubleParser"));
    assertTrue(result.contains("Schubfach"));
    assertTrue(result.contains("Boost Software License, Version 1.0"));
  }

  @Test
  void licenseTextContainsThirdPartySectionWhenDependencyLicensesPresent() {
    InputStream own =
        new ByteArrayInputStream(
            "Mozilla Public License Version 2.0\n".getBytes(StandardCharsets.UTF_8));
    InputStream notice = new ByteArrayInputStream("NOTICE info\n".getBytes(StandardCharsets.UTF_8));
    InputStream apache =
        new ByteArrayInputStream("Apache License\n".getBytes(StandardCharsets.UTF_8));
    InputStream bsd2 =
        new ByteArrayInputStream("BSD 2-Clause License\n".getBytes(StandardCharsets.UTF_8));
    InputStream bsd3 = new ByteArrayInputStream("BSD License\n".getBytes(StandardCharsets.UTF_8));

    String result = ProductInfo.licenseText(own, notice, apache, bsd2, bsd3);

    assertTrue(result.contains("Mozilla Public License Version 2.0"));
    assertTrue(result.contains("Third-party notices and licenses:"));
    assertTrue(result.contains("NOTICE info"));
    assertTrue(result.contains("Apache License"));
    assertTrue(result.contains("BSD 2-Clause License"));
    assertTrue(result.contains("BSD License"));
  }

  @Test
  void licenseTextReturnsFallbackWhenAllResourcesAbsent() {
    String result = ProductInfo.licenseText(null, null, null, null, null);

    assertFalse(result.isBlank());
    assertTrue(result.contains("not available"));
  }

  @Test
  void licenseTextAllowsThirdPartyOnlyBundlesAndUnreadableStreams() {
    InputStream apache =
        new ByteArrayInputStream("Apache License\n".getBytes(StandardCharsets.UTF_8));

    String thirdPartyOnly = ProductInfo.licenseText(null, null, apache, null, null);

    assertFalse(thirdPartyOnly.contains("Third-party notices and licenses:"));
    assertTrue(thirdPartyOnly.contains("Apache License"));
    assertTrue(
        ProductInfo.licenseText(new ThrowingInputStream(), null, null, null, null)
            .contains("not available"));
  }

  @Test
  void propertyFallbacksIgnoreUnreadableStreamsAndAnchorLicenseFallbackRemainsCallable() {
    assertEquals("unknown", ProductInfo.versionFrom(null, new ThrowingInputStream()));
    assertEquals("GridGrind", ProductInfo.descriptionFrom(new ThrowingInputStream()));
    assertFalse(ProductInfo.licenseText(Object.class).isBlank());
  }

  @Test
  void licenseTextAddsOneTrailingNewlineWhenSourceTextOmitsIt() {
    String text =
        ProductInfo.licenseText(
            new ByteArrayInputStream(
                "Mozilla Public License Version 2.0".getBytes(StandardCharsets.UTF_8)),
            null,
            null,
            null,
            null);

    assertEquals("Mozilla Public License Version 2.0\n", text);
  }

  @Test
  void helpFlagsPrintDedicatedSurfacesAndReturnExitCodeZero() throws IOException {
    String overview = assertHelpInvocationPrints("help");
    String protocol = assertHelpInvocationPrints("help", "protocol");
    String guidance = assertHelpInvocationPrints("help", "guidance");

    assertDefaultHelpMatchesOverview(overview);
    assertOverviewHelpSurface(overview);
    assertProtocolHelpSurface(protocol);
    assertGuidanceHelpSurface(guidance);
  }

  @Test
  void structuredIdentitySurfacesExposeMachineReadableReports() throws IOException {
    CliHelpReport helpReport =
        GridGrindCliJson.readBytes(
            runStructuredIdentitySurface("help", "--format", "structured"), CliHelpReport.class);
    CliVersionReport versionReport =
        GridGrindCliJson.readBytes(
            runStructuredIdentitySurface("version", "--format", "structured"),
            CliVersionReport.class);
    CliLicenseReport licenseReport =
        GridGrindCliJson.readBytes(
            runStructuredIdentitySurface("license", "--format", "structured"),
            CliLicenseReport.class);

    assertEquals("OVERVIEW", helpReport.topic());
    assertEquals(ProductInfo.version(), helpReport.version());
    assertEquals(ProductInfo.description(), helpReport.description());
    assertEquals(ProductInfo.documentRef(ProductInfo.version()), helpReport.documentRef());
    assertEquals(
        ProductInfo.containerImageRef(ProductInfo.version()), helpReport.containerImageRef());
    assertEquals(HelpSurface.CLI_SURFACE, helpReport.surface());

    assertEquals(ProductInfo.version(), versionReport.version());
    assertEquals(ProductInfo.description(), versionReport.description());
    assertEquals(ProductInfo.documentRef(ProductInfo.version()), versionReport.documentRef());
    assertEquals(
        ProductInfo.containerImageRef(ProductInfo.version()), versionReport.containerImageRef());

    assertEquals(ProductInfo.version(), licenseReport.version());
    assertFalse(licenseReport.licenseText().isBlank());
  }

  @Test
  void structuredIdentitySurfacesHonorPrettyFlag() throws IOException {
    String payload =
        new String(
            runStructuredIdentitySurface("version", "--format", "structured", "--pretty"),
            StandardCharsets.UTF_8);

    assertTrue(payload.startsWith("{\n"));
    assertTrue(payload.contains("\n  \"version\" : "));
  }

  private static void assertDefaultHelpMatchesOverview(String overview) throws IOException {
    ByteArrayOutputStream shortStdout = new ByteArrayOutputStream();
    int shortExitCode =
        new GridGrindCli().run(new String[0], new ByteArrayInputStream(new byte[0]), shortStdout);

    assertEquals(0, shortExitCode);
    assertEquals(overview, shortStdout.toString(StandardCharsets.UTF_8));
  }

  private static void assertOverviewHelpSurface(String overview) {
    String normalized = overview.replaceAll("\\s+", " ");
    assertTrue(overview.contains("Usage:"));
    assertTrue(overview.contains("recipe list | show <ID>"));
    assertTrue(overview.contains("schema request|grant|response"));
    assertTrue(overview.contains("help protocol"));
    assertTrue(overview.contains("--execution-root"));
    assertTrue(normalized.contains("payload there unchanged"));
    assertTrue(normalized.contains("never moves a primary payload to stderr"));
    assertTrue(overview.contains("docs/CLI.md"));
    assertFalse(overview.contains("--print-"));
  }

  private static void assertProtocolHelpSurface(String protocol) {
    String normalized = protocol.replaceAll("\\s+", " ");
    assertTrue(protocol.contains("Options:"));
    assertTrue(protocol.contains("Request:"));
    assertTrue(protocol.contains("Limits:"));
    assertTrue(protocol.contains("formulaEnvironment"));
    assertTrue(protocol.contains("EVALUATE_TARGETS"));
    assertTrue(protocol.contains("strategy.cells[]"));
    assertTrue(protocol.contains("stepId"));
    assertTrue(normalized.contains("Without --response"));
    assertTrue(normalized.contains("never moves a primary payload to stderr"));
    assertFalse(protocol.contains("Docker Example:"));
    assertFalse(protocol.contains("--doctor-request"));
  }

  private static void assertGuidanceHelpSurface(String guidance) {
    assertTrue(guidance.contains("Workflows:"));
    assertTrue(guidance.contains("Standard input:"));
    assertTrue(guidance.contains("run --request -"));
    assertTrue(guidance.contains("Docker Example:"));
    assertFalse(guidance.contains("{{CONTAINER_TAG}}"));
    assertTrue(guidance.contains("Discovery:"));
    assertTrue(guidance.contains("advisory:"));
    assertTrue(guidance.contains("recipe materialize"));
    assertTrue(guidance.contains("catalog"));
    assertFalse(guidance.contains("--print-"));
  }

  private static byte[] runStructuredIdentitySurface(String... args) throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode = new GridGrindCli().run(args, new ByteArrayInputStream(new byte[0]), stdout);
    assertEquals(0, exitCode);
    return stdout.toByteArray();
  }

  @Test
  void helpSurfacesUseVersionedAndFallbackReferencesCorrectly() {
    String unknownOverview = ProductInfo.helpText(CliCommand.HelpTopic.OVERVIEW, "unknown");
    assertTrue(unknownOverview.contains("blob/main/docs/CLI.md"));

    String overview = ProductInfo.helpText(CliCommand.HelpTopic.OVERVIEW, "0.9.0");
    String protocol = ProductInfo.helpText(CliCommand.HelpTopic.PROTOCOL, "0.9.0");
    String guidance = ProductInfo.helpText(CliCommand.HelpTopic.GUIDANCE, "0.9.0");

    assertTrue(overview.contains("blob/v0.9.0/docs/CLI.md"));
    assertTrue(overview.contains("blob/v0.9.0/docs/EXECUTION_MODEL.md"));
    assertTrue(overview.contains("blob/v0.9.0/docs/ERRORS.md"));
    assertTrue(protocol.contains("GridGrind 0.9.0"));
    assertFalse(protocol.contains("blob/v0.9.0/docs/CLI.md"));
    assertTrue(guidance.contains("ghcr.io/resoltico/gridgrind:0.9.0"));
    assertTrue(guidance.contains("blob/v0.9.0/docs/CLI.md"));
    assertTrue(guidance.contains("blob/v0.9.0/docs/EXECUTION_MODEL.md"));
    assertTrue(guidance.contains("blob/v0.9.0/docs/ERRORS.md"));
  }

  @Test
  void helpCanWriteToAnExplicitResponsePath() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-help-", ".txt");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"help", "guidance", "--response", responsePath.toString()},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    assertEquals(0, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertTrue(Files.readString(responsePath).contains("Workflows:"));
  }

  @Test
  void versionResponseFileCollisionPreservesTheOriginalVersionPayload() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-version-", ".json");
    Files.writeString(responsePath, "sentinel\n");
    ByteArrayOutputStream expected = new ByteArrayOutputStream();
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    GridGrindCli cli = new GridGrindCli();

    int directExitCode =
        cli.run(new String[] {"version"}, new ByteArrayInputStream(new byte[0]), expected);

    int exitCode =
        cli.run(
            new String[] {"version", "--response", responsePath.toString()},
            new ByteArrayInputStream(new byte[0]),
            stdout,
            stderr);

    dev.erst.gridgrind.cli.render.CliTransportNotice transportNotice =
        GridGrindCliJson.readBytes(
            stderr.toByteArray(), dev.erst.gridgrind.cli.render.CliTransportNotice.class);

    assertEquals(0, directExitCode);
    assertEquals(1, exitCode);
    assertEquals(
        expected.toString(StandardCharsets.UTF_8), stdout.toString(StandardCharsets.UTF_8));
    assertEquals(
        dev.erst.gridgrind.cli.render.CliTransportNotice.Destination.STDOUT,
        transportNotice.wroteTo());
    assertEquals(
        java.util.Optional.of(responsePath.toAbsolutePath().toString()),
        transportNotice.responsePath());
    assertEquals(
        "sentinel\n", Files.readString(responsePath), "existing response file must stay untouched");
  }

  @Test
  void guidanceHelpListsBuiltInExamplesAndRequiredPaths() {
    String help = ProductInfo.helpText(CliCommand.HelpTopic.GUIDANCE, "1.0.0");

    for (ShippedExampleEntry example : ShippedExamples.catalog().examples()) {
      assertTrue(help.contains(example.id()), () -> "missing example id " + example.id());
      assertTrue(
          help.contains(example.requestFileName()),
          () -> "missing example file name " + example.requestFileName());
      for (String requiredPath : example.requiredWorkspacePaths()) {
        assertTrue(help.contains(requiredPath), () -> "missing required path " + requiredPath);
      }
    }
  }

  @Test
  void protocolAndGuidanceBlocksDoNotMasqueradeAsOneAnother() {
    String protocol = ProductInfo.helpText(CliCommand.HelpTopic.PROTOCOL, "1.0.0");
    String guidance = ProductInfo.helpText(CliCommand.HelpTopic.GUIDANCE, "1.0.0");

    assertFalse(protocol.contains("Workflows:"));
    assertFalse(protocol.contains("docker run --rm -i"));
    assertTrue(guidance.contains("Workflows:"));
    assertTrue(guidance.contains("docker run --rm -i"));
  }

  @Test
  void guidanceHelpDocumentsDockerWorkdirUsage() {
    String version = "1.0.0";
    String help = ProductInfo.helpText(CliCommand.HelpTopic.GUIDANCE, version);
    String normalizedHelp = help.replace("\\", "").replaceAll("\\s+", " ");
    String expectedCommand =
        ContainerHelp.dockerMountedWorkdirExecutionCommand(ProductInfo.containerImageRef(version))
            .replaceAll("\\s+", " ");

    assertTrue(help.contains(ContainerHelp.dockerMountedWorkdirUserArgument()));
    assertTrue(help.contains(ContainerHelp.dockerMountedWorkdirVolumeArgument()));
    assertTrue(
        normalizedHelp.contains(expectedCommand),
        "guidance help must keep the mounted-directory Docker command in canonical order");
    assertFalse(help.contains("-w /workdir"));
    assertTrue(help.contains("--request request.json"));
    assertTrue(help.contains("--response response.json"));
  }

  @Test
  void productHeaderFormatsVersionAndDescription() {
    assertEquals(
        "GridGrind 1.0.0\nA description", ProductInfo.productHeader("1.0.0", "A description"));
  }

  private String assertHelpInvocationPrints(String... args) throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode = new GridGrindCli().run(args, new ByteArrayInputStream(new byte[0]), stdout);

    assertEquals(0, exitCode);
    return stdout.toString(StandardCharsets.UTF_8);
  }

  private static Path locateRepoRoot() {
    Path current = Path.of("").toAbsolutePath().normalize();
    while (current != null && !Files.exists(current.resolve("gradle.properties"))) {
      current = current.getParent();
    }
    if (current == null) {
      throw new AssertionError("test must run inside the GridGrind repository");
    }
    return current;
  }

  /**
   * Input stream that fails on every read so fallback branches stay covered without I/O fixtures.
   */
  private static final class ThrowingInputStream extends InputStream {
    @Override
    public int read() throws IOException {
      throw new IOException("boom");
    }

    @Override
    public int read(byte[] buffer, int offset, int length) throws IOException {
      throw new IOException("boom");
    }
  }
}
