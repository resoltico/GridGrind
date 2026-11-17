package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.nio.file.Path;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.Set;
import org.junit.jupiter.api.Test;
import org.opentest4j.AssertionFailedError;
import tools.jackson.databind.node.ObjectNode;

/** Keeps complete transport output and authored identities visible in golden comparisons. */
class GoldenSnapshotsTest {
  private static final String AUTHORED_ID = "{957F9ACA-DB10-4A22-9D21-0C5548B19B40}";
  private static final String GENERATED_ID = "{01234567-89ab-cdef-0123-456789abcdef}";

  @Test
  void trailingJsonAndPlaintextRemainObservable() {
    var multiple = GoldenSnapshots.observation(new GoldenLauncher.Result(0, "{}\n{}\n", ""));
    assertEquals(2, multiple.path("stdout").size());
    String contaminated = "{}\nunexpected native log\n";
    var plaintext = GoldenSnapshots.observation(new GoldenLauncher.Result(0, contaminated, ""));
    assertEquals(contaminated, plaintext.path("stdout").asString());
    assertNotEquals(GoldenSnapshots.observation(new GoldenLauncher.Result(0, "{}", "")), multiple);
  }

  @Test
  void launcherOverridesOnlyExactStderrAndRequireEveryKnownLauncher() {
    ObjectNode common = GoldenSnapshots.observation(new GoldenLauncher.Result(0, "{}", ""));
    ObjectNode stderr = GoldenSnapshots.JSON.createObjectNode().put("IN_PROCESS", "");
    stderr.putObject("SHADOW_JAR").put("type", "NATIVE_DIAGNOSTIC");
    assertEquals(
        common, GoldenSnapshots.expectedForLauncher(common, stderr, GoldenLauncher.IN_PROCESS));
    var shadow = GoldenSnapshots.expectedForLauncher(common, stderr, GoldenLauncher.SHADOW_JAR);
    assertEquals(common.path("stdout"), shadow.path("stdout"));
    assertEquals(common.path("exit"), shadow.path("exit"));
    assertEquals(stderr.path("SHADOW_JAR"), shadow.path("stderr"));
    assertNotEquals(shadow, GoldenSnapshots.observation(new GoldenLauncher.Result(0, "{}", "")));
    stderr.put("UNRECOGNIZED", "");
    assertThrows(
        AssertionFailedError.class,
        () -> GoldenSnapshots.expectedForLauncher(common, stderr, GoldenLauncher.SHADOW_JAR));
  }

  @Test
  void onlyCertifiedGeneratedSignatureSetupIdentitiesAreNormalized() {
    ObjectNode source = signature(AUTHORED_ID);
    ObjectNode copy = signature(GENERATED_ID);
    Path workspace = Path.of(".").toAbsolutePath();
    assertNotEquals(
        GoldenSnapshots.canonical(source, workspace), GoldenSnapshots.canonical(copy, workspace));
    String masked = GoldenSnapshots.canonical(copy, workspace, Set.of(), Set.of(GENERATED_ID));
    assertEquals(
        "<setupId-0>",
        GoldenSnapshots.JSON.readTree(masked).path("setup").path("setupId").asString());
    assertEquals(
        AUTHORED_ID,
        GoldenSnapshots.JSON
            .readTree(GoldenSnapshots.canonical(source, workspace, Set.of(), Set.of(GENERATED_ID)))
            .path("setup")
            .path("setupId")
            .asString());
    ObjectNode unrelated = GoldenSnapshots.JSON.createObjectNode().put("setupId", GENERATED_ID);
    assertEquals(
        GENERATED_ID,
        GoldenSnapshots.JSON
            .readTree(
                GoldenSnapshots.canonical(unrelated, workspace, Set.of(), Set.of(GENERATED_ID)))
            .path("setupId")
            .asString());
  }

  @Test
  void nativeIdentityMustUseCanonicalGuidShape() {
    GoldenSignatureIdentities.requireGuid(GENERATED_ID);
    assertThrows(
        AssertionFailedError.class, () -> GoldenSignatureIdentities.requireGuid("{1-1-1-1-1}"));
    assertThrows(
        AssertionFailedError.class,
        () -> GoldenSignatureIdentities.requireGuid(GENERATED_ID.substring(1)));
  }

  @Test
  void copyCertificationRejectsSharedIdentityOrSharedNativePart() {
    var source =
        new GoldenSignatureIdentities.SheetIdentity(
            "xl/drawings/source.vml", Map.of("Signature", AUTHORED_ID));
    var copy =
        new GoldenSignatureIdentities.SheetIdentity(
            "xl/drawings/copy.vml", Map.of("Signature", GENERATED_ID));
    Set<String> generated = new LinkedHashSet<>();
    GoldenSignatureIdentities.certifyCopy(source, copy, Set.of(AUTHORED_ID), generated);
    assertEquals(Set.of(GENERATED_ID), generated);
    var reusedId = new GoldenSignatureIdentities.SheetIdentity(copy.vmlPart(), source.ids());
    assertThrows(
        AssertionFailedError.class,
        () ->
            GoldenSignatureIdentities.certifyCopy(
                source, reusedId, Set.of(AUTHORED_ID), new LinkedHashSet<>()));
    var reusedPart = new GoldenSignatureIdentities.SheetIdentity(source.vmlPart(), copy.ids());
    assertThrows(
        AssertionFailedError.class,
        () ->
            GoldenSignatureIdentities.certifyCopy(
                source, reusedPart, Set.of(AUTHORED_ID), new LinkedHashSet<>()));
  }

  private static ObjectNode signature(String id) {
    ObjectNode signature = GoldenSnapshots.JSON.createObjectNode().put("type", "SIGNATURE_LINE");
    signature.putObject("setup").put("setupId", id);
    return signature;
  }
}
