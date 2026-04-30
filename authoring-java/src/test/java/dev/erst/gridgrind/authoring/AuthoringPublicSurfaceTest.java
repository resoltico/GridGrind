package dev.erst.gridgrind.authoring;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.module.ModuleDescriptor;
import java.lang.module.ModuleFinder;
import java.lang.reflect.Constructor;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.nio.file.Path;
import java.util.List;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Contract tests for the public Java authoring boundary. */
class AuthoringPublicSurfaceTest {
  private static final Set<String> BLOCKED_PUBLIC_TYPES =
      Set.of(
          "dev.erst.gridgrind.executor.",
          "dev.erst.gridgrind.contract.dto.CellInput",
          "dev.erst.gridgrind.contract.dto.ChartInput",
          "dev.erst.gridgrind.contract.dto.CommentInput",
          "dev.erst.gridgrind.contract.dto.DataValidationInput",
          "dev.erst.gridgrind.contract.dto.GridGrindResponse",
          "dev.erst.gridgrind.contract.dto.HyperlinkTarget",
          "dev.erst.gridgrind.contract.dto.NamedRangeScope",
          "dev.erst.gridgrind.contract.dto.NamedRangeTarget",
          "dev.erst.gridgrind.contract.dto.PivotTableInput",
          "dev.erst.gridgrind.contract.dto.PrintLayoutInput",
          "dev.erst.gridgrind.contract.dto.SheetPresentationInput",
          "dev.erst.gridgrind.contract.dto.TableInput",
          "dev.erst.gridgrind.contract.dto.WorkbookProtectionInput",
          "dev.erst.gridgrind.contract.assertion.ExpectedCellValue",
          "dev.erst.gridgrind.contract.source.BinarySourceInput",
          "dev.erst.gridgrind.contract.source.TextSourceInput");

  @Test
  void compiledAuthoringModuleNoLongerRequiresExecutorTransitively() throws Exception {
    Path compiledOutput =
        Path.of(GridGrindPlan.class.getProtectionDomain().getCodeSource().getLocation().toURI());
    ModuleDescriptor descriptor =
        ModuleFinder.of(compiledOutput)
            .find("dev.erst.gridgrind.authoring")
            .orElseThrow(() -> new AssertionError("compiled authoring module descriptor not found"))
            .descriptor();
    assertNotNull(descriptor);
    List<String> requiredModules =
        descriptor.requires().stream().map(ModuleDescriptor.Requires::name).toList();

    assertTrue(requiredModules.contains("dev.erst.gridgrind.contract"));
    assertFalse(requiredModules.contains("dev.erst.gridgrind.executor"));
  }

  @Test
  void gridGrindPlanStopsAtTheCanonicalBoundary() {
    boolean exposesRun =
        java.util.Arrays.stream(GridGrindPlan.class.getMethods())
            .map(Method::getName)
            .anyMatch("run"::equals);
    assertFalse(exposesRun);
  }

  @Test
  void checksAndQueriesAreNoLongerPublicApiSurface() throws Exception {
    assertFalse(
        Modifier.isPublic(Class.forName("dev.erst.gridgrind.authoring.Checks").getModifiers()));
    assertFalse(
        Modifier.isPublic(Class.forName("dev.erst.gridgrind.authoring.Queries").getModifiers()));
  }

  @Test
  void exportedPublicTypesDoNotExposeExecutorOrRawComplexContractTypes() {
    for (Class<?> type : exportedPublicTypes()) {
      for (Method method : type.getMethods()) {
        if (method.getDeclaringClass().equals(Object.class)) {
          continue;
        }
        assertAllowedType(
            method.getReturnType(), type.getName() + "#" + method.getName() + " return");
        for (Class<?> parameterType : method.getParameterTypes()) {
          assertAllowedType(parameterType, type.getName() + "#" + method.getName() + " parameter");
        }
      }
      for (Constructor<?> constructor : type.getConstructors()) {
        for (Class<?> parameterType : constructor.getParameterTypes()) {
          assertAllowedType(parameterType, type.getName() + " constructor");
        }
      }
    }
  }

  private static List<Class<?>> exportedPublicTypes() {
    return List.of(
        GridGrindPlan.class,
        Links.class,
        Links.Url.class,
        Links.Email.class,
        Links.FileTarget.class,
        Links.DocumentTarget.class,
        PlannedAssertion.class,
        PlannedInspection.class,
        PlannedMutation.class,
        Tables.class,
        Tables.Definition.class,
        Tables.NoStyle.class,
        Tables.NamedStyle.class,
        Targets.class,
        WorkbookTarget.class,
        SheetTarget.class,
        CellTarget.class,
        RangeTarget.class,
        WindowTarget.class,
        TableTarget.class,
        TableRowTarget.class,
        TableCellTarget.class,
        NamedRangeTarget.class,
        ChartTarget.class,
        PivotTableTarget.class,
        Values.class,
        Values.Blank.class,
        Values.Text.class,
        Values.NumericValue.class,
        Values.BooleanValue.class,
        Values.DateValue.class,
        Values.DateTimeValue.class,
        Values.Formula.class,
        Values.Comment.class,
        Values.ExpectedBlank.class,
        Values.ExpectedText.class,
        Values.ExpectedNumber.class,
        Values.ExpectedBoolean.class,
        Values.ExpectedError.class,
        Values.InlineText.class,
        Values.Utf8FileText.class,
        Values.StandardInputText.class);
  }

  private static void assertAllowedType(Class<?> type, String location) {
    String typeName = type.getName();
    for (String blocked : BLOCKED_PUBLIC_TYPES) {
      assertFalse(
          typeName.startsWith(blocked), () -> location + " exposes blocked type " + typeName);
    }
  }
}
