package dev.erst.gridgrind.core.authoring;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;

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
          "dev.erst.gridgrind.engine.api.",
          "dev.erst.gridgrind.engine.runtime.",
          "dev.erst.gridgrind.core.model.value.CellValueSpec",
          "dev.erst.gridgrind.core.model.chart.ChartSpec",
          "dev.erst.gridgrind.core.model.drawing.CommentSpec",
          "dev.erst.gridgrind.core.model.validation.DataValidationSpec",
          "dev.erst.gridgrind.core.result.WorkbookResult",
          "dev.erst.gridgrind.core.model.value.HyperlinkTarget",
          "dev.erst.gridgrind.core.vocabulary.NamedRangeScope",
          "dev.erst.gridgrind.core.model.formula.NamedRangeTarget",
          "dev.erst.gridgrind.core.model.pivot.PivotTableSpec",
          "dev.erst.gridgrind.core.model.print.PrintLayoutSpec",
          "dev.erst.gridgrind.core.model.layout.SheetPresentationSpec",
          "dev.erst.gridgrind.core.model.table.TableSpec",
          "dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec",
          "dev.erst.gridgrind.core.source.BinarySourceSpec",
          "dev.erst.gridgrind.core.source.TextSourceSpec");

  @Test
  void coreAuthoringDoesNotRequireEngineOrPoi() throws Exception {
    Path compiledOutput =
        Path.of(GridGrindPlan.class.getProtectionDomain().getCodeSource().getLocation().toURI());
    ModuleDescriptor descriptor =
        ModuleFinder.of(compiledOutput)
            .find("dev.erst.gridgrind.core")
            .orElseThrow(() -> new AssertionError("compiled core module descriptor not found"))
            .descriptor();
    assertNotNull(descriptor);
    List<String> requiredModules =
        descriptor.requires().stream().map(ModuleDescriptor.Requires::name).toList();

    assertFalse(requiredModules.stream().anyMatch(name -> name.startsWith("org.apache.poi")));
    assertFalse(requiredModules.contains("dev.erst.gridgrind.engine"));
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
  void internalChecksAndQueryFamiliesStayOutOfThePublicApiSurface() throws Exception {
    for (String typeName :
        List.of(
            "dev.erst.gridgrind.core.authoring.Checks",
            "dev.erst.gridgrind.core.authoring.WorkbookQueries",
            "dev.erst.gridgrind.core.authoring.SheetQueries",
            "dev.erst.gridgrind.core.authoring.WorkbookAssetQueries",
            "dev.erst.gridgrind.core.authoring.InspectionSurfaceQueries",
            "dev.erst.gridgrind.core.authoring.InspectionAnalysisQueries")) {
      assertFalse(Modifier.isPublic(Class.forName(typeName).getModifiers()));
    }
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
        ExpectedValues.class,
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
