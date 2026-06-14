package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.gradle.api.GradleException;
import org.junit.jupiter.api.Test;

class JavaSourceShapePolicyTest {
  @Test
  void loadsReviewedExactSurfaceMetadata() throws IOException {
    Path policyFile =
        writePolicy(
            """
            EXACT\tengine/src/main/java/dev/erst/gridgrind/excel/ExcelSheetCells.java\tsheet-cell-surface\t124\t18\t17\t6\t4\t2\t2\t4\tworkbook-core-api\tCHECK\t2026-08-15\tSplit read preview from write authoring before adding another unrelated cell interaction family.\tExcelSheetCells remains the singular sheet-level cell surface.
            DEFAULT\t*\tproduction-source\t260\t22\t12\t28\t8\t2\t4\t8\trepo-shape\tCHECK\t-\t-\tOrdinary production Java sources should stay small.
            """);

    JavaSourceShapePolicy policy = JavaSourceShapePolicy.load(policyFile);

    JavaSourceShapePolicy.Rule reviewedRule = policy.rules().getFirst();
    assertEquals(JavaSourceShapePolicy.MatchKind.EXACT, reviewedRule.kind());
    assertEquals("2026-08-15", reviewedRule.reviewExpiresOn().toString());
    assertEquals(2, reviewedRule.maxNestedTypes());
    assertEquals(
        "Split read preview from write authoring before adding another unrelated cell interaction family.",
        reviewedRule.splitTrigger());

    JavaSourceShapePolicy.Rule defaultRule = policy.rules().getLast();
    assertNull(defaultRule.reviewExpiresOn());
    assertNull(defaultRule.splitTrigger());
  }

  @Test
  void rejectsExactSurfaceWithoutReviewMetadata() throws IOException {
    Path policyFile =
        writePolicy(
            """
            EXACT\tengine/src/main/java/dev/erst/gridgrind/excel/ExcelSheetCells.java\tsheet-cell-surface\t124\t18\t17\t6\t4\t2\t2\t4\tworkbook-core-api\tCHECK\t-\t-\tExcelSheetCells remains the singular sheet-level cell surface.
            DEFAULT\t*\tproduction-source\t260\t22\t12\t28\t8\t2\t4\t8\trepo-shape\tCHECK\t-\t-\tOrdinary production Java sources should stay small.
            """);

    assertThrows(GradleException.class, () -> JavaSourceShapePolicy.load(policyFile));
  }

  @Test
  void rejectsBroadRuleThatPretendsToBeReviewed() throws IOException {
    Path policyFile =
        writePolicy(
            """
            PREFIX\tcli/src/main/java/dev/erst/gridgrind/cli/\tcli-support\t500\t28\t18\t30\t8\t3\t10\t16\tcli\tCHECK\t2026-08-31\tSplit command parsing from discovery commands.\tCLI orchestration aggregates command wiring and help surfaces.
            DEFAULT\t*\tproduction-source\t260\t22\t12\t28\t8\t2\t4\t8\trepo-shape\tCHECK\t-\t-\tOrdinary production Java sources should stay small.
            """);

    JavaSourceShapePolicy policy = JavaSourceShapePolicy.load(policyFile);
    assertEquals(JavaSourceShapePolicy.MatchKind.PREFIX, policy.rules().getFirst().kind());
    assertEquals(
        "2026-08-31", policy.rules().getFirst().reviewExpiresOn().toString());
  }

  @Test
  void rejectsBroadRuleThatLoosensDefaultWithoutReviewMetadata() throws IOException {
    Path policyFile =
        writePolicy(
            """
            PREFIX\tcli/src/main/java/dev/erst/gridgrind/cli/\tcli-support\t500\t28\t18\t30\t8\t3\t10\t16\tcli\tCHECK\t-\t-\tCLI orchestration aggregates command wiring and help surfaces.
            DEFAULT\t*\tproduction-source\t260\t22\t12\t28\t8\t2\t4\t8\trepo-shape\tCHECK\t-\t-\tOrdinary production Java sources should stay small.
            """);

    assertThrows(GradleException.class, () -> JavaSourceShapePolicy.load(policyFile));
  }

  private static Path writePolicy(String contents) throws IOException {
    Path policyFile = Files.createTempFile("gridgrind-source-shape-policy", ".tsv");
    Files.writeString(policyFile, contents);
    return policyFile;
  }
}
