package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

class JavaSourceShapeAnalyzerTest {
  @Test
  void countsImplicitInterfaceSurfaceAndNestedTypes() throws IOException {
    Path sourceFile = Files.createTempFile("gridgrind-shape-analyzer", ".java");
    Files.writeString(
        sourceFile,
        """
        interface Surface {
          void implicit();
          default void defaultMethod() {}
          static void utility() {}
          private void hidden() {}
        }

        @interface Marker {
          String value();
        }

        class Holder {
          void packagePrivate() {}
          public void visible() {}
          private void hidden() {}

          interface Nested {
            void nestedImplicit();
          }
        }
        """);

    JavaSourceShapeAnalyzer.Metrics metrics = new JavaSourceShapeAnalyzer().analyze(sourceFile, 26);

    assertEquals(20L, metrics.lineCount());
    assertEquals(0, metrics.importCount());
    assertEquals(3, metrics.topLevelTypeCount());
    assertEquals(1, metrics.nestedTypeCount());
    assertEquals(9, metrics.methodCount());
    assertEquals(6, metrics.publicMethodCount());
    assertEquals(0, metrics.fieldCount());
    assertEquals(0, metrics.switchCount());
    assertEquals(0, metrics.maxSwitchArms());
  }
}
