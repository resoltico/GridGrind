package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/** Build-failing audit of the permanent ArchUnit dependency, execution, and policy contract. */
class ArchUnitIntegrationAuditTest {
  private static final List<String> ARCHITECTURE_SOURCE_ROOTS =
      List.of(
          "executor/src/main/java/dev/erst/gridgrind/architecture",
          "executor/src/architectureTest/java");
  private static final List<String> FORBIDDEN_BASELINE_APIS =
      List.of(
          "FreezingArchRule",
          "ignoreDependency(",
          "allowEmptyShould(",
          "setFailOnEmptyShould(",
          "ArchIgnore",
          "@Disabled");

  @Test
  void versionCatalogPinsOnlyTheJUnit6Integration() throws IOException {
    String versionCatalog = read("gradle/libs.versions.toml");

    assertTrue(
        versionCatalog.contains("archunit = \"1.5.0\""),
        "the version catalog must pin the reviewed ArchUnit release");
    assertTrue(
        versionCatalog.contains(
            "archunit-junit6 = { module = \"com.tngtech.archunit:archunit-junit6\", version.ref = \"archunit\" }"),
        "the version catalog must expose the JUnit 6 convenience integration");
    assertFalse(
        versionCatalog.contains("archunit-junit5"),
        "the JUnit 5 integration must not enter the JUnit 6 build");
  }

  @Test
  void executorOwnsAnAlignedAndAnalyzedArchitectureSourceSet() throws IOException {
    String executorBuild = read("executor/build.gradle.kts");

    assertTrue(
        executorBuild.contains("sourceSets.create(\"architectureTest\")"),
        "executor must own the dedicated architecture source set");
    assertTrue(
        executorBuild.contains("runtimeClasspath += mainSourceSet.runtimeClasspath"),
        "architecture tests must consume the main runtime classpath without a duplicate compile view");
    assertTrue(
        executorBuild.contains("platform(libs.junit.bom)"),
        "architecture tests must use the repository JUnit BOM");
    assertTrue(
        executorBuild.contains("libs.archunit.junit6"),
        "architecture tests must use ArchUnit's JUnit 6 integration");
    assertTrue(
        executorBuild.contains("libs.junit.jupiter"),
        "deliberate-violation regressions must run on Jupiter");
    assertTrue(
        executorBuild.contains("libs.junit.platform.launcher"),
        "architecture tests must pin the matching Platform Launcher");
    assertTrue(
        executorBuild.contains("includeEngines(\"archunit\")"),
        "the architecture-rule task must fail when the ArchUnit engine is unavailable");
    assertTrue(
        executorBuild.contains("includeEngines(\"junit-jupiter\")"),
        "the architecture runtime-contract task must execute on Jupiter independently");
    assertTrue(
        executorBuild.contains("failOnNoDiscoveredTests = true"),
        "both architecture tasks must reject an empty discovered-test population");
    assertTrue(
        executorBuild.contains("tasks.register<Test>(\"architectureTestContract\")"),
        "the effective architecture runtime contract must have its own Jupiter task");
    assertTrue(
        executorBuild.contains("dependsOn(architectureTest)"),
        "executor check must execute architecture tests");
    assertTrue(
        executorBuild.contains("pmdArchitectureTest"),
        "architecture-test source must remain under PMD");
    assertFalse(
        executorBuild.contains("junit.testFilter"),
        "the architecture task must not filter out registered rules");
    assertFalse(
        executorBuild.contains("ignoreFailures"),
        "the architecture task and PMD must remain blocking");
  }

  @Test
  void architectureSuiteImportsOnlyProductLocationsAndChecksItsEffectiveRuntimePolicy()
      throws IOException {
    String suite =
        read(
            "executor/src/architectureTest/java/dev/erst/gridgrind/architecture/ProductArchitectureTest.java");
    String runtimeContract =
        read(
            "executor/src/architectureTest/java/dev/erst/gridgrind/architecture/ArchitectureGateContractTest.java");

    assertTrue(
        suite.contains("locations = GridGrindProductLocations.class"),
        "the rule suite must import explicit compiled product locations rather than the full test runtime");
    assertTrue(
        suite.contains("ImportOption.DoNotIncludeGradleTestFixtures.class"),
        "the rule suite must reject Gradle test-fixture locations defensively");
    assertTrue(
        suite.contains("ArchTests.in(ProductArchitectureRules.class)"),
        "the rule suite must execute the complete declared rule inventory");
    assertTrue(
        runtimeContract.contains(
            "ArchConfiguration.get().getProperty(\"archRule.failOnEmptyShould\")"),
        "the architecture runtime contract must assert the effective fail-closed configuration");
    assertTrue(
        runtimeContract.contains(
            "assertEquals(13, ProductArchitectureRules.mandatoryRules().size())"),
        "the architecture runtime contract must ratchet the mandatory rule inventory");
  }

  @Test
  void rootCheckExposesAndExecutesTheArchitectureAggregate() throws IOException {
    String rootConventions =
        read(
            "gradle/build-logic/src/main/kotlin/dev/erst/gridgrind/buildlogic/GridGrindRootConventionsPlugin.kt");
    String architectureVerification =
        read(
            "gradle/build-logic/src/main/kotlin/dev/erst/gridgrind/buildlogic/GridGrindArchitectureVerification.kt");

    assertTrue(
        rootConventions.contains("registerGridGrindArchitectureCheck()"),
        "root conventions must register the architecture aggregate");
    assertTrue(
        rootConventions.contains("checkTask.dependsOn(architectureCheck)"),
        "root check must depend on the architecture aggregate");
    assertTrue(
        architectureVerification.contains("tasks.register(\"architectureCheck\""),
        "the aggregate must remain a discoverable verification task");
    assertTrue(
        architectureVerification.contains(
            "architectureTask.dependsOn(\":executor:architectureTest\")"),
        "the aggregate must delegate to the executor architecture suite");
    assertTrue(
        architectureVerification.contains(
            "architectureTask.dependsOn(\":executor:architectureTestContract\")"),
        "the aggregate must verify the architecture task's effective runtime contract");
    assertTrue(
        architectureVerification.contains(
            "architectureTask.dependsOn(\":executor:pmdArchitectureTest\")"),
        "the aggregate must analyze architecture-test source quality");
  }

  @Test
  void guidanceDescribesThePermanentFailClosedGate() throws IOException {
    String agentsExtra = read(".codex/AGENTS_EXTRA.md");
    String developerGuide = read("docs/DEVELOPER.md");
    String developerGradle = read("docs/DEVELOPER_GRADLE.md");

    assertTrue(
        agentsExtra.contains("ArchUnit") && agentsExtra.contains("FreezingArchRule"),
        "agent guidance must teach the executable no-baseline policy");
    assertTrue(
        developerGuide.contains("ArchUnit") && developerGuide.contains("bytecode-level"),
        "the developer guide must describe bytecode architecture enforcement");
    assertTrue(
        developerGradle.contains("architectureTest"),
        "the Gradle guide must publish the architecture source set");
    assertTrue(
        developerGradle.contains("archunit-junit6"),
        "the Gradle guide must publish the JUnit 6 integration");
    assertTrue(
        developerGradle.contains("permanent generally available repository gate"),
        "the Gradle guide must describe ArchUnit as generally available");
  }

  @Test
  void rulesFailClosedWithoutFrozenOrIgnoredViolations() throws IOException {
    assertEquals(
        "archRule.failOnEmptyShould=true",
        read("executor/src/architectureTest/resources/archunit.properties").strip());

    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    for (Path sourcePath : architectureSources(repositoryRoot)) {
      String source = Files.readString(sourcePath);
      String relativePath = repositoryRoot.relativize(sourcePath).toString();
      for (String forbiddenApi : FORBIDDEN_BASELINE_APIS) {
        assertFalse(
            source.contains(forbiddenApi), () -> relativePath + " must not use " + forbiddenApi);
      }
    }
  }

  private static List<Path> architectureSources(Path repositoryRoot) throws IOException {
    List<Path> sources = new ArrayList<>();
    for (String sourceRoot : ARCHITECTURE_SOURCE_ROOTS) {
      try (Stream<Path> paths = Files.walk(repositoryRoot.resolve(sourceRoot))) {
        sources.addAll(
            paths
                .filter(path -> Files.isRegularFile(path) && path.toString().endsWith(".java"))
                .toList());
      }
    }
    return sources.stream().sorted().toList();
  }

  private static String read(String relativePath) throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    return Files.readString(repositoryRoot.resolve(relativePath));
  }
}
