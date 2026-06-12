package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.gradle.api.GradleException;
import org.junit.jupiter.api.Test;

class JavaSemanticShapePolicyTest {
  @Test
  void loadsReviewedExactAndPrefixRules() throws IOException {
    Path policyFile =
        writePolicy(
            """
            PREFIX\tengine/src/main/java/dev/erst/gridgrind/excel/\texcel-support\tCouplingBetweenObjects,GodClass\tworkbook-core\t2026-08-31\tBreak workbook-core helpers by document concern before broadening responsibilities.\tWorkbook-core semantic-shape debt remains under active decomposition.
            EXACT\tengine/src/main/java/dev/erst/gridgrind/excel/ExcelSheet.java\texcel-sheet-facade\tCouplingBetweenObjects\tworkbook-core-api\t2026-08-15\tSplit the facade into narrower public sub-surfaces before adding new capability families.\tExcelSheet is the public workbook-sheet facade.
            """);

    JavaSemanticShapePolicy policy = JavaSemanticShapePolicy.load(policyFile);

    JavaSemanticShapePolicy.Rule exact =
        policy.matchingRule("engine/src/main/java/dev/erst/gridgrind/excel/ExcelSheet.java");
    assertEquals(JavaSourceShapePolicy.MatchKind.EXACT, exact.kind());
    assertTrue(exact.allows("CouplingBetweenObjects"));

    JavaSemanticShapePolicy.Rule prefix =
        policy.matchingRule(
            "engine/src/main/java/dev/erst/gridgrind/excel/ExcelChartSourceSupport.java");
    assertEquals(JavaSourceShapePolicy.MatchKind.PREFIX, prefix.kind());
    assertTrue(prefix.allows("GodClass"));
  }

  @Test
  void rejectsDefaultRules() throws IOException {
    Path policyFile =
        writePolicy(
            """
            DEFAULT\t*\tsemantic-default\tGodClass\trepo-shape\t2026-08-31\tRemove the fake default.\tDefaults are not allowed here.
            """);

    assertThrows(GradleException.class, () -> JavaSemanticShapePolicy.load(policyFile));
  }

  @Test
  void rejectsRulesWithoutExpiry() throws IOException {
    Path policyFile =
        writePolicy(
            """
            EXACT\tcontract/src/main/java/dev/erst/gridgrind/contract/json/GridGrindJsonProblemMessageSupport.java\tcontract-json\tGodClass\tcontract-json\t-\tSplit message rendering by concern.\tJSON error wording debt is under review.
            """);

    assertThrows(GradleException.class, () -> JavaSemanticShapePolicy.load(policyFile));
  }

  private static Path writePolicy(String contents) throws IOException {
    Path policyFile = Files.createTempFile("gridgrind-semantic-shape-policy", ".tsv");
    Files.writeString(policyFile, contents);
    return policyFile;
  }
}
