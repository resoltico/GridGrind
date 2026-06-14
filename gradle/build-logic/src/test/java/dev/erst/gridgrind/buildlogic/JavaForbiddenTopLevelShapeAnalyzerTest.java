package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

class JavaForbiddenTopLevelShapeAnalyzerTest {
  @Test
  void flagsTaggedUnionRecordsAcrossTopLevelAndNestedVariantSurfaces() throws IOException {
    Path sourceFile = Files.createTempFile("gridgrind-forbidden-shape", ".java");
    Files.writeString(
        sourceFile,
        """
        import java.util.Optional;

        record Tagged(String type, Optional<String> left, Optional<String> right) {}

        record StatusTagged(String status, Optional<String> left, Optional<String> right) {}

        record Wide(
            Optional<String> a1,
            Optional<String> a2,
            Optional<String> a3,
            Optional<String> a4,
            Optional<String> a5,
            Optional<String> a6,
            String a7,
            String a8,
            String a9,
            String a10,
            String a11,
            String a12,
            String a13) {}

        sealed interface SafeUnion permits SafeUnion.Left, SafeUnion.Right {
          record Left(Optional<String> value) implements SafeUnion {}
          record Right(Optional<String> value) implements SafeUnion {}
        }

        record PropertyBearing(String shape, Optional<String> left, Optional<String> right) {}

        sealed interface UnsafeUnion permits UnsafeUnion.Payload {
          record Payload(String status, Optional<String> left, Optional<String> right)
              implements UnsafeUnion {}
        }
        """);

    List<JavaForbiddenTopLevelShapeAnalyzer.Violation> violations =
        new JavaForbiddenTopLevelShapeAnalyzer().analyze(sourceFile, 26);

    assertEquals(
        3, violations.size(), () -> "Expected the two direct violations plus the nested variant.");
    assertTrue(
        violations.stream()
            .anyMatch(
                violation ->
                    violation.typeName().equals("StatusTagged")
                        && violation.reason().contains("Record mixes discriminator slots")));
    assertTrue(
        violations.stream()
            .anyMatch(
                violation ->
                    violation
                        .reason()
                        .contains("Record combines 13 state slots with 6 Optional-bearing components")));
    assertTrue(
        violations.stream()
            .anyMatch(
                violation ->
                    violation.typeName().equals("UnsafeUnion.Payload")
                        && violation.reason().contains("Record mixes discriminator slots")));
    assertTrue(
        violations.stream().noneMatch(violation -> violation.typeName().equals("PropertyBearing")));
    assertTrue(violations.stream().noneMatch(violation -> violation.typeName().equals("Tagged")));
  }
}
