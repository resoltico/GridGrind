package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.WorkbookCommand;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Regression coverage for the extracted Jazzer value-factory seam. */
class OperationSequenceValueFactoryTest {
  /** Parameter kinds used to synthesize representative inputs for factory coverage. */
  private enum CoverageParameterKind {
    DATA,
    BOOLEAN,
    INT,
    STRING,
    PATH
  }

  /** One reflective-free invocation of a single value-factory method under test. */
  @FunctionalInterface
  private interface CoverageInvocation {
    Object invoke(Object[] arguments) throws IOException;
  }

  /** One bounded value-factory coverage case with its input shape and assertion contract. */
  private record CoverageCase(
      String name,
      List<CoverageParameterKind> parameterKinds,
      boolean returnsVoid,
      CoverageInvocation invocation) {
    CoverageCase {
      parameterKinds = List.copyOf(parameterKinds);
    }

    int parameterCount() {
      return parameterKinds.size();
    }

    Object invoke(Object[] arguments) throws IOException {
      return invocation.invoke(arguments);
    }
  }

  private static final int SELECTOR_VARIANTS = 16;
  private static final int PAYLOAD_BYTES = 512;

  @TempDir Path tempDir;

  @Test
  void factoryMethodsProduceBoundedValuesAcrossSelectorSweep() throws IOException {
    for (CoverageCase coverageCase : coverageCases()) {
      for (int selector = 0; selector < SELECTOR_VARIANTS; selector++) {
        Object[] arguments = argumentsFor(coverageCase, selector);
        Object result = coverageCase.invoke(arguments);
        assertMethodOutcome(coverageCase, arguments, result);
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

  private static List<CoverageCase> coverageCases() {
    return List.of(
        new CoverageCase(
            "nextPaneInput",
            List.of(CoverageParameterKind.DATA),
            false,
            args -> OperationSequenceValueFactory.nextPaneInput((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextExcelSheetPane",
            List.of(CoverageParameterKind.DATA),
            false,
            args -> OperationSequenceValueFactory.nextExcelSheetPane((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextPrintLayoutInput",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextPrintLayoutInput((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextExcelPrintLayout",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelPrintLayout((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextProtocolPaneRegion",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextProtocolPaneRegion((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextPaneRegion",
            List.of(CoverageParameterKind.DATA),
            false,
            args -> OperationSequenceValueFactory.nextPaneRegion((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextSheetCopyPosition",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextSheetCopyPosition((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextExcelSheetCopyPosition",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelSheetCopyPosition(
                    (GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextSelectedSheetNames",
            List.of(
                CoverageParameterKind.DATA,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING),
            false,
            args ->
                OperationSequenceValueFactory.nextSelectedSheetNames(
                    (GridGrindFuzzData) args[0], (String) args[1], (String) args[2])),
        new CoverageCase(
            "nextProtocolSheetVisibility",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextProtocolSheetVisibility(
                    (GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextSheetVisibility",
            List.of(CoverageParameterKind.DATA),
            false,
            args -> OperationSequenceValueFactory.nextSheetVisibility((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextProtocolSheetProtectionSettings",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextProtocolSheetProtectionSettings(
                    (GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextSheetProtectionSettings",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextSheetProtectionSettings(
                    (GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextWorkflowStorage",
            List.of(
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextWorkflowStorage(
                    (String) args[0], (String) args[1], (GridGrindFuzzData) args[2])),
        new CoverageCase(
            "writeExistingWorkbook",
            List.of(
                CoverageParameterKind.PATH,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.DATA),
            true,
            args -> {
              OperationSequenceValueFactory.writeExistingWorkbook(
                  (Path) args[0], (String) args[1], (String) args[2], (GridGrindFuzzData) args[3]);
              return null;
            }),
        new CoverageCase(
            "nextHyperlinkTarget",
            List.of(CoverageParameterKind.DATA),
            false,
            args -> OperationSequenceValueFactory.nextHyperlinkTarget((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextExcelHyperlink",
            List.of(CoverageParameterKind.DATA),
            false,
            args -> OperationSequenceValueFactory.nextExcelHyperlink((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextCommentInput",
            List.of(CoverageParameterKind.DATA),
            false,
            args -> OperationSequenceValueFactory.nextCommentInput((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextPictureInput",
            List.of(CoverageParameterKind.DATA),
            false,
            args -> OperationSequenceValueFactory.nextPictureInput((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextChartInput",
            List.of(CoverageParameterKind.DATA),
            false,
            args -> OperationSequenceValueFactory.nextChartInput((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextShapeInput",
            List.of(CoverageParameterKind.DATA),
            false,
            args -> OperationSequenceValueFactory.nextShapeInput((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextEmbeddedObjectInput",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextEmbeddedObjectInput((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextDrawingAnchorInput",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextDrawingAnchorInput((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextPictureDataInput",
            List.of(),
            false,
            args -> OperationSequenceValueFactory.nextPictureDataInput()),
        new CoverageCase(
            "nextExcelPictureDefinition",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelPictureDefinition(
                    (GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextExcelChartDefinition",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelChartDefinition(
                    (GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextExcelShapeDefinition",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelShapeDefinition(
                    (GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextExcelEmbeddedObjectDefinition",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelEmbeddedObjectDefinition(
                    (GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextExcelDrawingAnchor",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelDrawingAnchor((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextDrawingAnchorBehavior",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextDrawingAnchorBehavior(
                    (GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextDrawingObjectName",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextDrawingObjectName((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextDrawingBinaryObjectName",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextDrawingBinaryObjectName(
                    (GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextDataValidationInput",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextDataValidationInput((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextExcelDataValidationDefinition",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelDataValidationDefinition(
                    (GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextConditionalFormattingInput",
            List.of(CoverageParameterKind.DATA, CoverageParameterKind.BOOLEAN),
            false,
            args ->
                OperationSequenceValueFactory.nextConditionalFormattingInput(
                    (GridGrindFuzzData) args[0], (boolean) args[1])),
        new CoverageCase(
            "nextExcelConditionalFormattingBlockDefinition",
            List.of(CoverageParameterKind.DATA, CoverageParameterKind.BOOLEAN),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelConditionalFormattingBlockDefinition(
                    (GridGrindFuzzData) args[0], (boolean) args[1])),
        new CoverageCase(
            "nextDifferentialStyleInput",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextDifferentialStyleInput(
                    (GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextExcelDifferentialStyle",
            List.of(CoverageParameterKind.DATA),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelDifferentialStyle(
                    (GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextRangeSelector",
            List.of(
                CoverageParameterKind.DATA,
                CoverageParameterKind.STRING,
                CoverageParameterKind.BOOLEAN),
            false,
            args ->
                OperationSequenceValueFactory.nextRangeSelector(
                    (GridGrindFuzzData) args[0], (String) args[1], (boolean) args[2])),
        new CoverageCase(
            "nextExcelRangeSelection",
            List.of(CoverageParameterKind.DATA, CoverageParameterKind.BOOLEAN),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelRangeSelection(
                    (GridGrindFuzzData) args[0], (boolean) args[1])),
        new CoverageCase(
            "nextAutofilterRange",
            List.of(CoverageParameterKind.BOOLEAN),
            false,
            args -> OperationSequenceValueFactory.nextAutofilterRange((boolean) args[0])),
        new CoverageCase(
            "nextCopySheetName",
            List.of(CoverageParameterKind.STRING),
            false,
            args -> OperationSequenceValueFactory.nextCopySheetName((String) args[0])),
        new CoverageCase(
            "nextTableInput",
            List.of(
                CoverageParameterKind.DATA,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.BOOLEAN),
            false,
            args ->
                OperationSequenceValueFactory.nextTableInput(
                    (GridGrindFuzzData) args[0],
                    (String) args[1],
                    (String) args[2],
                    (boolean) args[3])),
        new CoverageCase(
            "nextExcelTableDefinition",
            List.of(
                CoverageParameterKind.DATA,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.BOOLEAN),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelTableDefinition(
                    (GridGrindFuzzData) args[0],
                    (String) args[1],
                    (String) args[2],
                    (boolean) args[3])),
        new CoverageCase(
            "nextTableStyleInput",
            List.of(CoverageParameterKind.DATA),
            false,
            args -> OperationSequenceValueFactory.nextTableStyleInput((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextExcelTableStyle",
            List.of(CoverageParameterKind.DATA),
            false,
            args -> OperationSequenceValueFactory.nextExcelTableStyle((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextTableSelector",
            List.of(
                CoverageParameterKind.DATA,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING),
            false,
            args ->
                OperationSequenceValueFactory.nextTableSelector(
                    (GridGrindFuzzData) args[0], (String) args[1], (String) args[2])),
        new CoverageCase(
            "nextPivotTableSelector",
            List.of(
                CoverageParameterKind.DATA,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.BOOLEAN),
            false,
            args ->
                OperationSequenceValueFactory.nextPivotTableSelector(
                    (GridGrindFuzzData) args[0],
                    (String) args[1],
                    (String) args[2],
                    (String) args[3],
                    (boolean) args[4])),
        new CoverageCase(
            "nextPivotTableInput",
            List.of(
                CoverageParameterKind.DATA,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.BOOLEAN,
                CoverageParameterKind.BOOLEAN),
            false,
            args ->
                OperationSequenceValueFactory.nextPivotTableInput(
                    (GridGrindFuzzData) args[0],
                    (String) args[1],
                    (String) args[2],
                    (String) args[3],
                    (String) args[4],
                    (boolean) args[5],
                    (boolean) args[6])),
        new CoverageCase(
            "nextPivotTableSource",
            List.of(
                CoverageParameterKind.DATA,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.BOOLEAN,
                CoverageParameterKind.BOOLEAN),
            false,
            args ->
                OperationSequenceValueFactory.nextPivotTableSource(
                    (GridGrindFuzzData) args[0],
                    (String) args[1],
                    (String) args[2],
                    (String) args[3],
                    (boolean) args[4],
                    (boolean) args[5])),
        new CoverageCase(
            "nextExcelPivotTableDefinition",
            List.of(
                CoverageParameterKind.DATA,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.BOOLEAN,
                CoverageParameterKind.BOOLEAN),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelPivotTableDefinition(
                    (GridGrindFuzzData) args[0],
                    (String) args[1],
                    (String) args[2],
                    (String) args[3],
                    (String) args[4],
                    (boolean) args[5],
                    (boolean) args[6])),
        new CoverageCase(
            "nextExcelPivotTableSource",
            List.of(
                CoverageParameterKind.DATA,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.STRING,
                CoverageParameterKind.BOOLEAN,
                CoverageParameterKind.BOOLEAN),
            false,
            args ->
                OperationSequenceValueFactory.nextExcelPivotTableSource(
                    (GridGrindFuzzData) args[0],
                    (String) args[1],
                    (String) args[2],
                    (String) args[3],
                    (boolean) args[4],
                    (boolean) args[5])),
        new CoverageCase(
            "nextTableName",
            List.of(
                CoverageParameterKind.DATA,
                CoverageParameterKind.BOOLEAN,
                CoverageParameterKind.STRING),
            false,
            args ->
                OperationSequenceValueFactory.nextTableName(
                    (GridGrindFuzzData) args[0], (boolean) args[1], (String) args[2])),
        new CoverageCase(
            "nextExcelComment",
            List.of(CoverageParameterKind.DATA),
            false,
            args -> OperationSequenceValueFactory.nextExcelComment((GridGrindFuzzData) args[0])),
        new CoverageCase(
            "nextNamedRangeName",
            List.of(CoverageParameterKind.DATA, CoverageParameterKind.BOOLEAN),
            false,
            args ->
                OperationSequenceValueFactory.nextNamedRangeName(
                    (GridGrindFuzzData) args[0], (boolean) args[1])),
        new CoverageCase(
            "nextPivotTableName",
            List.of(CoverageParameterKind.DATA, CoverageParameterKind.BOOLEAN),
            false,
            args ->
                OperationSequenceValueFactory.nextPivotTableName(
                    (GridGrindFuzzData) args[0], (boolean) args[1])));
  }

  private Object[] argumentsFor(CoverageCase coverageCase, int selector) {
    Object[] arguments = new Object[coverageCase.parameterCount()];
    int booleanIndex = 0;
    int intIndex = 0;
    int stringIndex = 0;
    for (int index = 0; index < coverageCase.parameterCount(); index++) {
      CoverageParameterKind parameterKind = coverageCase.parameterKinds().get(index);
      switch (parameterKind) {
        case DATA -> arguments[index] = GridGrindFuzzData.replay(selectorBytes(selector));
        case BOOLEAN -> {
          arguments[index] = booleanArgument(coverageCase, booleanIndex, selector);
          booleanIndex++;
        }
        case INT -> {
          arguments[index] = intIndex == 0 ? 0 : 10;
          intIndex++;
        }
        case STRING -> {
          arguments[index] = stringArgument(stringIndex);
          stringIndex++;
        }
        case PATH ->
            arguments[index] = tempDir.resolve(coverageCase.name() + "-" + selector + ".xlsx");
      }
    }
    return arguments;
  }

  private void assertMethodOutcome(CoverageCase coverageCase, Object[] arguments, Object result)
      throws IOException {
    if (coverageCase.returnsVoid()) {
      assertWrittenWorkbook((Path) arguments[0]);
      return;
    }

    assertNotNull(result, coverageCase.name() + " should not return null");
    if ("nextCopySheetName".equals(coverageCase.name())) {
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

  private static boolean booleanArgument(
      CoverageCase coverageCase, int booleanIndex, int selector) {
    return switch (coverageCase.name()) {
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
