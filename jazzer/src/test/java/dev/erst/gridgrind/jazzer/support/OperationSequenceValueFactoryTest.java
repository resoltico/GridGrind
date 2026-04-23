package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.WorkbookCommand;
import java.io.IOException;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Regression coverage for the extracted Jazzer value-factory seam. */
class OperationSequenceValueFactoryTest {
  private static final int SELECTOR_VARIANTS = 16;
  private static final int PAYLOAD_BYTES = 512;

  @TempDir Path tempDir;

  @Test
  void factoryMethodsProduceBoundedValuesAcrossSelectorSweep()
      throws IOException, ReflectiveOperationException {
    for (Method method : factoryMethods()) {
      for (int selector = 0; selector < SELECTOR_VARIANTS; selector++) {
        Object[] arguments = argumentsFor(method, selector);
        Object result = invoke(method, arguments);
        assertMethodOutcome(method, arguments, result);
      }
    }
  }

  @Test
  void operationSequenceModelBuildsAndCleansUpAcrossSelectorSweep() throws IOException {
    int successfulWorkflows = 0;
    int successfulCommandSequences = 0;
    for (int selector = 0; selector < SELECTOR_VARIANTS; selector++) {
      try {
        GeneratedProtocolWorkflow workflow =
            OperationSequenceModel.nextProtocolWorkflow(
                GridGrindFuzzData.replay(selectorBytes(selector)));
        successfulWorkflows++;
        assertNotNull(workflow.request());
        assertFalse(workflow.cleanupRoots().isEmpty());
        for (Path cleanupRoot : workflow.cleanupRoots()) {
          assertTrue(Files.isDirectory(cleanupRoot));
        }
        workflow.cleanup();
        for (Path cleanupRoot : workflow.cleanupRoots()) {
          assertFalse(Files.exists(cleanupRoot));
        }
      } catch (IllegalArgumentException expected) {
        // The protocol-workflow harness treats generated invalid requests as an expected outcome.
      }

      try {
        List<WorkbookCommand> commands =
            OperationSequenceModel.nextWorkbookCommands(
                GridGrindFuzzData.replay(selectorBytes(selector)));
        successfulCommandSequences++;
        assertFalse(commands.isEmpty());
      } catch (IllegalArgumentException expected) {
        // The workbook-command harness also treats generated invalid commands as expected.
      }
    }
    assertTrue(successfulWorkflows > 0);
    assertTrue(successfulCommandSequences > 0);
  }

  private static List<Method> factoryMethods() {
    List<Method> methods = new ArrayList<>();
    for (Method method : OperationSequenceValueFactory.class.getDeclaredMethods()) {
      if (!Modifier.isStatic(method.getModifiers()) || Modifier.isPrivate(method.getModifiers())) {
        continue;
      }
      if (method.isSynthetic()) {
        continue;
      }
      methods.add(method);
    }
    methods.sort(Comparator.comparing(Method::getName));
    return methods;
  }

  private Object[] argumentsFor(Method method, int selector) {
    Object[] arguments = new Object[method.getParameterCount()];
    int booleanIndex = 0;
    int intIndex = 0;
    int stringIndex = 0;
    for (int index = 0; index < method.getParameterCount(); index++) {
      Class<?> parameterType = method.getParameterTypes()[index];
      if (parameterType == GridGrindFuzzData.class) {
        arguments[index] = GridGrindFuzzData.replay(selectorBytes(selector));
      } else if (parameterType == boolean.class) {
        arguments[index] = booleanArgument(method, booleanIndex, selector);
        booleanIndex++;
      } else if (parameterType == int.class) {
        arguments[index] = intIndex == 0 ? 0 : 10;
        intIndex++;
      } else if (parameterType == String.class) {
        arguments[index] = stringArgument(stringIndex);
        stringIndex++;
      } else if (parameterType == Path.class) {
        arguments[index] = tempDir.resolve(method.getName() + "-" + selector + ".xlsx");
      } else {
        throw new IllegalArgumentException(
            "Unsupported parameter type " + parameterType.getName() + " on " + method.getName());
      }
    }
    return arguments;
  }

  private static Object invoke(Method method, Object[] arguments)
      throws ReflectiveOperationException {
    return method.invoke(null, arguments);
  }

  private void assertMethodOutcome(Method method, Object[] arguments, Object result)
      throws IOException {
    if ("nextOptionalInt".equals(method.getName())) {
      if (result != null) {
        int candidate = (Integer) result;
        assertTrue(candidate >= 0);
        assertTrue(candidate <= 10);
      }
      return;
    }
    if (method.getReturnType() == void.class) {
      assertWrittenWorkbook((Path) arguments[0]);
      return;
    }

    assertNotNull(result, method.getName() + " should not return null");
    if ("nextCopySheetName".equals(method.getName())) {
      assertTrue(((String) result).length() <= 31);
    }
    if (result instanceof OperationSequenceValueFactory.WorkflowStorage storage) {
      assertTrue(Files.isDirectory(storage.cleanupRoot()));
      if (storage.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile) {
        assertWrittenWorkbook(Path.of(existingFile.path()));
      }
      deleteRecursively(storage.cleanupRoot());
      assertFalse(Files.exists(storage.cleanupRoot()));
    }
  }

  private static byte[] selectorBytes(int selector) {
    byte[] input = new byte[PAYLOAD_BYTES];
    Arrays.fill(input, (byte) selector);
    return input;
  }

  private static boolean booleanArgument(Method method, int booleanIndex, int selector) {
    return switch (method.getName()) {
      case "nextAutofilterRange",
          "nextExcelPivotTableDefinition",
          "nextExcelPivotTableSource",
          "nextExcelRangeSelection",
          "nextExcelTableDefinition",
          "nextPivotTableInput",
          "nextPivotTableSelector",
          "nextPivotTableSource",
          "nextRangeSelector",
          "nextTableInput",
          "nextTableName" ->
          true;
      default -> ((selector + booleanIndex) & 1) == 1;
    };
  }

  private static String stringArgument(int index) {
    return switch (index) {
      case 0 -> "Budget";
      case 1 -> "Archive";
      case 2 -> "BudgetTotal";
      case 3 -> "OpsTable";
      default -> "OpsPivot";
    };
  }

  private static void assertWrittenWorkbook(Path path) {
    assertTrue(Files.isRegularFile(path));
    assertDoesNotThrow(
        () -> {
          try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(path))) {
            assertTrue(workbook.getNumberOfSheets() > 0);
          }
        });
  }

  private static void deleteRecursively(Path root) throws IOException {
    if (!Files.exists(root)) {
      return;
    }
    try (var paths = Files.walk(root)) {
      for (Path path : paths.sorted(Comparator.reverseOrder()).toList()) {
        Files.deleteIfExists(path);
      }
    }
  }
}
