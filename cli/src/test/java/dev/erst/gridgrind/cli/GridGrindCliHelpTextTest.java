package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.CliHelpReport;
import dev.erst.gridgrind.cli.discovery.CliLicenseReport;
import dev.erst.gridgrind.cli.discovery.CliVersionReport;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.cli.discovery.ShippedExampleEntry;
import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.catalog.GridGrindContainerRuntimeText;
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

    assertEquals(properties.getProperty("version"), GridGrindCliProductInfo.version());
    assertEquals(
        properties.getProperty("gridgrindDescription"), GridGrindCliProductInfo.description());
  }

  @Test
  void versionFromReturnsBundledVersionWhenImplementationVersionIsAbsent() throws IOException {
    assertEquals(GridGrindCliProductInfo.version(), GridGrindCli.versionFrom(null));
  }

  @Test
  void versionFromReturnsUnknownWhenBundledVersionIsUnavailable() {
    assertEquals("unknown", GridGrindCliProductInfo.versionFrom(null, Object.class));
  }

  @Test
  void versionFromIgnoresBlankImplementationVersionAndBlankBundledVersion() {
    assertEquals(
        "unknown",
        GridGrindCliProductInfo.versionFrom(
            "   ", new ByteArrayInputStream("version=\n".getBytes(StandardCharsets.UTF_8))));
  }

  @Test
  void versionFromReturnsVersionWhenImplementationVersionIsPresent() {
    assertEquals("0.4.1", GridGrindCli.versionFrom("0.4.1"));
  }

  @Test
  void versionFlagPrintsVersionLineToStdoutAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--version"}, new ByteArrayInputStream(new byte[0]), stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8);
    assertTrue(output.startsWith("GridGrind " + GridGrindCliProductInfo.version() + "\n"));
    assertTrue(output.endsWith("\n"));
    assertTrue(output.lines().count() >= 2);
  }

  @Test
  void licenseFlagPrintsLicenseTextToStdoutAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--license"}, new ByteArrayInputStream(new byte[0]), stdout);

    assertEquals(0, exitCode);
    assertFalse(stdout.toString(StandardCharsets.UTF_8).isBlank());
  }

  @Test
  void descriptionFromReturnsFallbackWhenPropertiesAreUnavailableOrBlank() throws IOException {
    assertEquals("GridGrind", GridGrindCli.descriptionFrom(Object.class));
    assertEquals(
        "GridGrind",
        GridGrindCli.descriptionFrom(
            new ByteArrayInputStream("description=\n".getBytes(StandardCharsets.UTF_8))));
  }

  @Test
  void licenseTextContainsMitLicenseWhenResourcePresent() {
    String mit = "MIT License\n\nCopyright (c) 2026 Ervins Strauhmanis\n";
    InputStream own = new ByteArrayInputStream(mit.getBytes(StandardCharsets.UTF_8));

    String result = GridGrindCli.licenseText(own, null, null, null, null);

    assertTrue(result.contains("MIT License"));
    assertTrue(result.contains("Ervins Strauhmanis"));
    assertFalse(result.contains("Third-party notices and licenses:"));
  }

  @Test
  void packagedLicenseTextCoversResolvedRuntimeTerms() {
    String result = GridGrindCli.licenseText(GridGrindCli.class);

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
    InputStream own = new ByteArrayInputStream("MIT License\n".getBytes(StandardCharsets.UTF_8));
    InputStream notice = new ByteArrayInputStream("NOTICE info\n".getBytes(StandardCharsets.UTF_8));
    InputStream apache =
        new ByteArrayInputStream("Apache License\n".getBytes(StandardCharsets.UTF_8));
    InputStream bsd2 =
        new ByteArrayInputStream("BSD 2-Clause License\n".getBytes(StandardCharsets.UTF_8));
    InputStream bsd3 = new ByteArrayInputStream("BSD License\n".getBytes(StandardCharsets.UTF_8));

    String result = GridGrindCli.licenseText(own, notice, apache, bsd2, bsd3);

    assertTrue(result.contains("MIT License"));
    assertTrue(result.contains("Third-party notices and licenses:"));
    assertTrue(result.contains("NOTICE info"));
    assertTrue(result.contains("Apache License"));
    assertTrue(result.contains("BSD 2-Clause License"));
    assertTrue(result.contains("BSD License"));
  }

  @Test
  void licenseTextReturnsFallbackWhenAllResourcesAbsent() {
    String result = GridGrindCli.licenseText(null, null, null, null, null);

    assertFalse(result.isBlank());
    assertTrue(result.contains("not available"));
  }

  @Test
  void licenseTextAllowsThirdPartyOnlyBundlesAndUnreadableStreams() {
    InputStream apache =
        new ByteArrayInputStream("Apache License\n".getBytes(StandardCharsets.UTF_8));

    String thirdPartyOnly = GridGrindCli.licenseText(null, null, apache, null, null);

    assertFalse(thirdPartyOnly.contains("Third-party notices and licenses:"));
    assertTrue(thirdPartyOnly.contains("Apache License"));
    assertTrue(
        GridGrindCli.licenseText(new ThrowingInputStream(), null, null, null, null)
            .contains("not available"));
  }

  @Test
  void propertyFallbacksIgnoreUnreadableStreamsAndAnchorLicenseFallbackRemainsCallable() {
    assertEquals("unknown", GridGrindCliProductInfo.versionFrom(null, new ThrowingInputStream()));
    assertEquals("GridGrind", GridGrindCli.descriptionFrom(new ThrowingInputStream()));
    assertFalse(GridGrindCli.licenseText(Object.class).isBlank());
  }

  @Test
  void licenseTextAddsOneTrailingNewlineWhenSourceTextOmitsIt() {
    String text =
        GridGrindCli.licenseText(
            new ByteArrayInputStream("MIT License".getBytes(StandardCharsets.UTF_8)),
            null,
            null,
            null,
            null);

    assertEquals("MIT License\n", text);
  }

  @Test
  void helpFlagsPrintDedicatedSurfacesAndReturnExitCodeZero() throws IOException {
    String overview = assertHelpInvocationPrints("--help");
    String protocol = assertHelpInvocationPrints("--help-protocol");
    String guidance = assertHelpInvocationPrints("--help-guidance");

    assertShortHelpAliasMatchesOverview(overview);
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
    assertEquals(GridGrindCliProductInfo.version(), helpReport.version());
    assertEquals(GridGrindCliProductInfo.description(), helpReport.description());
    assertEquals(
        GridGrindCliProductInfo.documentRef(GridGrindCliProductInfo.version()),
        helpReport.documentRef());
    assertEquals(
        GridGrindCliProductInfo.containerImageRef(GridGrindCliProductInfo.version()),
        helpReport.containerImageRef());
    assertEquals(GridGrindProtocolCatalogCliSurface.CLI_SURFACE, helpReport.surface());

    assertEquals(GridGrindCliProductInfo.version(), versionReport.version());
    assertEquals(GridGrindCliProductInfo.description(), versionReport.description());
    assertEquals(
        GridGrindCliProductInfo.documentRef(GridGrindCliProductInfo.version()),
        versionReport.documentRef());
    assertEquals(
        GridGrindCliProductInfo.containerImageRef(GridGrindCliProductInfo.version()),
        versionReport.containerImageRef());

    assertEquals(GridGrindCliProductInfo.version(), licenseReport.version());
    assertFalse(licenseReport.licenseText().isBlank());
  }

  @Test
  void structuredIdentitySurfacesHonorPrettyFlag() throws IOException {
    String payload =
        new String(
            runStructuredIdentitySurface("--version", "--format", "structured", "--pretty"),
            StandardCharsets.UTF_8);

    assertTrue(payload.startsWith("{\n"));
    assertTrue(payload.contains("\n  \"version\" : "));
  }

  private static void assertShortHelpAliasMatchesOverview(String overview) throws IOException {
    ByteArrayOutputStream shortStdout = new ByteArrayOutputStream();
    int shortExitCode =
        new GridGrindCli()
            .run(new String[] {"-h"}, new ByteArrayInputStream(new byte[0]), shortStdout);

    assertEquals(0, shortExitCode);
    assertEquals(overview, shortStdout.toString(StandardCharsets.UTF_8));
  }

  private static void assertOverviewHelpSurface(String overview) {
    String normalizedOverview = overview.replaceAll("\\s+", " ");
    assertTrue(overview.contains("Primary Commands:"));
    assertTrue(overview.contains("Command Rules:"));
    assertTrue(overview.contains("Next Commands:"));
    assertTrue(overview.contains("--print-recipe --lookup <id>"));
    assertTrue(overview.contains("--print-recipe-keyword-match --query <text>"));
    assertTrue(overview.contains("--print-protocol-catalog --lookup <lookup-id>"));
    assertFalse(overview.contains("--print-protocol-catalog --lookup <id>|<group>:<id>"));
    assertTrue(overview.contains("--print-protocol-catalog --search <text>"));
    assertTrue(overview.contains("--execution-root <path>"));
    assertTrue(overview.contains("--temp-root <path>"));
    assertTrue(overview.contains("--help, -h"));
    assertTrue(overview.contains("--version"));
    assertTrue(overview.contains("--license"));
    assertTrue(overview.contains("--print-recipe-catalog [--lookup <id>]"));
    assertTrue(overview.contains("--help-protocol"));
    assertTrue(overview.contains("--help-guidance"));
    assertTrue(overview.contains("Use --format structured"));
    assertTrue(overview.contains("Use --pretty"));
    assertTrue(overview.contains("--pretty"));
    assertTrue(overview.contains("Without --response"));
    assertTrue(overview.contains("compact transport notice"));
    assertTrue(
        normalizedOverview.contains(
            "with writable stdout recovers that already-rendered payload there unchanged"));
    assertTrue(normalizedOverview.contains("never moves a primary payload to stderr"));
    assertTrue(overview.contains("docs/QUICK_REFERENCE.md"));
    assertFalse(overview.contains("Minimal Valid Request:"));
    assertFalse(overview.contains("Built-in generated examples:"));
  }

  private static void assertProtocolHelpSurface(String protocol) {
    String normalizedProtocol = protocol.replaceAll("\\s+", " ");
    assertTrue(protocol.contains("Authoritative Contract Scope:"));
    assertTrue(protocol.contains("Flags:"));
    assertTrue(protocol.contains("Limits:"));
    assertTrue(protocol.contains("Request:"));
    assertTrue(protocol.contains("File Workflow:"));
    assertTrue(protocol.contains("Coordinate Systems:"));
    assertFalse(protocol.contains("Minimal Valid Request:"));
    assertTrue(protocol.contains("--print-recipe --lookup <id>"));
    assertTrue(protocol.contains("--print-recipe-keyword-match --query <text>"));
    assertTrue(protocol.contains("--print-protocol-catalog --lookup <lookup-id>"));
    assertTrue(protocol.contains("--execution-root <path>"));
    assertTrue(protocol.contains("--temp-root <path>"));
    assertTrue(protocol.contains("--pretty"));
    assertFalse(protocol.contains("--lookup GET_CELLS"));
    assertFalse(protocol.contains("--lookup nestedTypes:cellInputTypes"));
    assertFalse(protocol.contains("summary-first"));
    assertTrue(protocol.contains("execution.mode is a typed discriminator"));
    assertTrue(
        normalizedProtocol.contains(
            "execution may include any subset of execution.mode, execution.journal, execution.calculation, and execution.assertionMode"));
    assertTrue(
        normalizedProtocol.contains(
            "The bare --print-protocol-catalog command emits only the compact index."));
    assertTrue(protocol.contains("projectedByFacets"));
    assertTrue(protocol.contains("noteRefs"));
    assertTrue(protocol.contains("enumValueDocs"));
    assertTrue(
        normalizedProtocol.contains("Scoped --lookup payloads may annotate conditional fields"));
    assertTrue(normalizedProtocol.contains("publish shared notes"));
    assertTrue(protocol.contains("STREAMING_WRITE"));
    assertTrue(protocol.contains("VERBOSE streams compact JSONL progress events to stderr"));
    assertTrue(protocol.contains("formulaEnvironment.missingWorkbookPolicy"));
    assertTrue(protocol.contains("USE_CACHED_VALUE"));
    assertTrue(protocol.contains("formulaEnvironment.udfToolpacks[]"));
    assertTrue(protocol.contains("GET_CELLS addresses"));
    assertTrue(protocol.contains("EVALUATE_TARGETS requires strategy.cells[]"));
    assertTrue(protocol.contains("stepId must be unique within steps[]"));
    assertTrue(
        normalizedProtocol.contains(
            "Without --response, the command payload is the sole stdout content"));
    assertTrue(
        normalizedProtocol.contains("With --response, that payload is written to the new file"));
    assertTrue(
        normalizedProtocol.contains(
            "response-file write failure recovers the already-rendered payload there unchanged and adds one compact transport notice"));
    assertTrue(normalizedProtocol.contains("never moves a primary payload to stderr"));
    assertFalse(protocol.contains("Workflows:"));
    assertFalse(protocol.contains("Docker Example:"));
  }

  private static void assertGuidanceHelpSurface(String guidance) {
    assertTrue(guidance.contains("Operator Guidance Scope:"));
    assertTrue(guidance.contains("Workflow Playbooks:"));
    assertTrue(guidance.contains("Stdin Example:"));
    assertTrue(
        guidance.contains("gridgrind --print-request-template | gridgrind --execution-root ."));
    assertTrue(guidance.contains("Docker Example:"));
    assertFalse(guidance.contains("{{CONTAINER_TAG}}"));
    assertTrue(guidance.contains("Discovery:"));
    assertTrue(guidance.contains("Unified recipe catalog entries:"));
    assertTrue(guidance.contains("advisory:"));
    assertTrue(guidance.contains("Print one recipe:"));
    assertTrue(guidance.contains("gridgrind --print-recipe --lookup"));
    assertTrue(
        guidance.contains("The bare --print-protocol-catalog output is intentionally compact:"));
    assertTrue(guidance.contains("Shared reusable notes such as request-owned path resolution"));
    assertFalse(guidance.contains("starter scaffolds"));
    assertFalse(guidance.contains("Authoritative Contract Scope:"));
    assertFalse(guidance.contains("Minimal Valid Request:"));
  }

  private static byte[] runStructuredIdentitySurface(String... args) throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode = new GridGrindCli().run(args, new ByteArrayInputStream(new byte[0]), stdout);
    assertEquals(0, exitCode);
    return stdout.toByteArray();
  }

  @Test
  void helpSurfacesUseVersionedAndFallbackReferencesCorrectly() {
    String unknownOverview = GridGrindCli.helpText("unknown");
    assertTrue(unknownOverview.contains("blob/main/docs/QUICK_REFERENCE.md"));

    String overview = GridGrindCli.helpText("0.9.0");
    String protocol = GridGrindCliProductInfo.helpText(CliCommand.HelpTopic.PROTOCOL, "0.9.0");
    String guidance = GridGrindCliProductInfo.helpText(CliCommand.HelpTopic.GUIDANCE, "0.9.0");

    assertTrue(overview.contains("blob/v0.9.0/docs/QUICK_REFERENCE.md"));
    assertFalse(overview.contains("blob/v0.9.0/docs/OPERATIONS.md"));
    assertFalse(overview.contains("blob/v0.9.0/docs/ERRORS.md"));
    assertTrue(protocol.contains("GridGrind 0.9.0"));
    assertFalse(protocol.contains("blob/v0.9.0/docs/QUICK_REFERENCE.md"));
    assertTrue(guidance.contains("ghcr.io/resoltico/gridgrind:0.9.0"));
    assertTrue(guidance.contains("blob/v0.9.0/docs/QUICK_REFERENCE.md"));
    assertTrue(guidance.contains("blob/v0.9.0/docs/OPERATIONS.md"));
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
                new String[] {"--help-guidance", "--response", responsePath.toString()},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    assertEquals(0, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertTrue(Files.readString(responsePath).contains("Workflow Playbooks:"));
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
        cli.run(new String[] {"--version"}, new ByteArrayInputStream(new byte[0]), expected);

    int exitCode =
        cli.run(
            new String[] {"--version", "--response", responsePath.toString()},
            new ByteArrayInputStream(new byte[0]),
            stdout,
            stderr);

    dev.erst.gridgrind.cli.discovery.CliTransportNotice transportNotice =
        GridGrindCliJson.readBytes(
            stderr.toByteArray(), dev.erst.gridgrind.cli.discovery.CliTransportNotice.class);

    assertEquals(0, directExitCode);
    assertEquals(1, exitCode);
    assertEquals(
        expected.toString(StandardCharsets.UTF_8), stdout.toString(StandardCharsets.UTF_8));
    assertEquals(
        dev.erst.gridgrind.cli.discovery.CliTransportNotice.Destination.STDOUT,
        transportNotice.wroteTo());
    assertEquals(
        java.util.Optional.of(responsePath.toAbsolutePath().toString()),
        transportNotice.responsePath());
    assertEquals(
        "sentinel\n", Files.readString(responsePath), "existing response file must stay untouched");
  }

  @Test
  void guidanceHelpListsBuiltInExamplesAndRequiredPaths() {
    String help = GridGrindCliProductInfo.helpText(CliCommand.HelpTopic.GUIDANCE, "1.0.0");

    for (ShippedExampleEntry example : GridGrindShippedExamples.catalog().examples()) {
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
    String protocol = GridGrindCliProductInfo.helpText(CliCommand.HelpTopic.PROTOCOL, "1.0.0");
    String guidance = GridGrindCliProductInfo.helpText(CliCommand.HelpTopic.GUIDANCE, "1.0.0");

    assertFalse(protocol.contains("Start from the minimal request:"));
    assertFalse(protocol.contains("docker run --rm -i"));
    assertTrue(guidance.contains("Start from the minimal request:"));
    assertTrue(guidance.contains("docker run --rm -i"));
  }

  @Test
  void guidanceHelpDocumentsDockerWorkdirUsage() {
    String version = "1.0.0";
    String help = GridGrindCliProductInfo.helpText(CliCommand.HelpTopic.GUIDANCE, version);
    String normalizedHelp = help.replace("\\", "").replaceAll("\\s+", " ");
    String expectedCommand =
        GridGrindContainerRuntimeText.dockerMountedWorkdirExecutionCommand(
                GridGrindCliProductInfo.containerImageRef(version))
            .replaceAll("\\s+", " ");

    assertTrue(help.contains(GridGrindContainerRuntimeText.dockerMountedWorkdirUserArgument()));
    assertTrue(help.contains(GridGrindContainerRuntimeText.dockerMountedWorkdirVolumeArgument()));
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
        "GridGrind 1.0.0\nA description", GridGrindCli.productHeader("1.0.0", "A description"));
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
