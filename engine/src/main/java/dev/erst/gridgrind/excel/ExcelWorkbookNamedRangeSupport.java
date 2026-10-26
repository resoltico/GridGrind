package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import org.apache.poi.ss.usermodel.Name;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedNames;

/** Named-range ownership and projection support for {@link ExcelWorkbook}. */
final class ExcelWorkbookNamedRangeSupport {
  private ExcelWorkbookNamedRangeSupport() {}

  static ExcelWorkbook setNamedRange(
      ExcelWorkbook workbook, StructuredMutationAction.SetNamedRange definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    int index =
        switch (definition.scope()) {
          case dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook _ -> -1;
          case dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet sheet ->
              workbook.requiredSheetIndex(sheet.sheetName().value());
        };
    String formula = ExcelNamedRangeTargets.toFormula(definition.target());
    Name name = existingName(workbook, definition.name().value(), index);
    if (name == null) name = workbook.context().workbook().createName();
    name.setSheetIndex(index);
    name.setNameName(definition.name().value());
    name.setRefersToFormula(formula);
    return workbook;
  }

  static ExcelWorkbook deleteNamedRange(
      ExcelWorkbook workbook, String name, NamedRangeScopeReport scope) {
    Name existingName = requiredName(workbook, name, scope);
    workbook.context().workbook().removeName(existingName);
    return workbook;
  }

  static List<NamedRangeReport> namedRanges(ExcelWorkbook workbook) {
    return namedRangesFromPoi(workbook);
  }

  /** Returns raw existing names that Apache POI did not model as {@link Name} instances. */
  static List<NamedRangeReport> unmodeledObservedNamedRanges(ExcelWorkbook workbook) {
    CTDefinedNames definedNames = workbook.context().workbook().getCTWorkbook().getDefinedNames();
    if (definedNames == null) {
      return List.of();
    }
    Set<NamedRangeIdentity> modeledNames = new LinkedHashSet<>();
    for (NamedRangeReport snapshot : namedRangesFromPoi(workbook)) {
      modeledNames.add(new NamedRangeIdentity(snapshot.name().value(), snapshot.scope()));
    }
    List<NamedRangeReport> namedRanges = new ArrayList<>();
    for (CTDefinedName definedName : definedNames.getDefinedNameList()) {
      String name = Objects.requireNonNullElse(definedName.getName(), "");
      if (name.isBlank()
          || !shouldExpose(name, definedName.getFunction(), definedName.getHidden())) {
        continue;
      }
      NamedRangeScopeReport scope = toScope(workbook, definedName);
      if (modeledNames.add(new NamedRangeIdentity(name, scope))) {
        namedRanges.add(
            snapshot(name, scope, Objects.requireNonNullElse(definedName.getStringValue(), "")));
      }
    }
    return List.copyOf(namedRanges);
  }

  private static List<NamedRangeReport> namedRangesFromPoi(ExcelWorkbook workbook) {
    List<NamedRangeReport> namedRanges = new ArrayList<>();
    for (Name name : workbook.context().workbook().getAllNames()) {
      if (!shouldExpose(name)) {
        continue;
      }
      namedRanges.add(
          snapshot(
              name.getNameName(),
              toScope(workbook, name),
              Objects.requireNonNullElse(name.getRefersToFormula(), "")));
    }
    return List.copyOf(namedRanges);
  }

  private static NamedRangeReport snapshot(
      String name, NamedRangeScopeReport scope, String refersToFormula) {
    var target = ExcelNamedRangeTargets.resolveTarget(refersToFormula, scope);
    if (target.isEmpty()) {
      return new NamedRangeReport.FormulaReport(
          new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel(name), scope, refersToFormula);
    }
    return new NamedRangeReport.RangeReport(
        new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel(name),
        scope,
        refersToFormula,
        target.orElseThrow());
  }

  static boolean scopeMatches(ExcelWorkbook workbook, Name candidate, NamedRangeScopeReport scope) {
    return switch (scope) {
      case NamedRangeScopeReport.Workbook _ -> candidate.getSheetIndex() < 0;
      case NamedRangeScopeReport.Sheet sheetScope ->
          candidate.getSheetIndex() == workbook.requiredSheetIndex(sheetScope.sheetName().value());
    };
  }

  static boolean shouldExpose(Name name) {
    return shouldExpose(name.getNameName(), name.isFunctionName(), name.isHidden());
  }

  static boolean shouldExpose(String nameName, boolean functionName, boolean hidden) {
    return !functionName
        && !hidden
        && nameName != null
        && !nameName.startsWith("_xlnm.")
        && !nameName.startsWith("_XLNM.");
  }

  private static Name requiredName(
      ExcelWorkbook workbook, String name, NamedRangeScopeReport scope) {
    Name existingName = existingName(workbook, name, scopeIndex(workbook, scope));
    if (existingName == null) {
      throw new NamedRangeNotFoundException(name, scope);
    }
    return existingName;
  }

  private static @Nullable Name existingName(ExcelWorkbook workbook, String name, int scopeIndex) {
    String validatedName =
        dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName(name);
    return workbook.context().workbook().getAllNames().stream()
        .filter(candidate -> candidate.getNameName().equalsIgnoreCase(validatedName))
        .filter(
            candidate ->
                scopeIndex < 0
                    ? candidate.getSheetIndex() < 0
                    : candidate.getSheetIndex() == scopeIndex)
        .findFirst()
        .orElse(null);
  }

  private static int scopeIndex(ExcelWorkbook workbook, NamedRangeScopeReport scope) {
    return switch (scope) {
      case NamedRangeScopeReport.Workbook _ -> -1;
      case NamedRangeScopeReport.Sheet sheet ->
          workbook.requiredSheetIndex(sheet.sheetName().value());
    };
  }

  private static NamedRangeScopeReport toScope(ExcelWorkbook workbook, Name name) {
    int sheetIndex = name.getSheetIndex();
    if (sheetIndex < 0) {
      return new NamedRangeScopeReport.Workbook();
    }
    return new NamedRangeScopeReport.Sheet(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(
            workbook.context().workbook().getSheetName(sheetIndex)));
  }

  private static NamedRangeScopeReport toScope(ExcelWorkbook workbook, CTDefinedName definedName) {
    if (!definedName.isSetLocalSheetId()) {
      return new NamedRangeScopeReport.Workbook();
    }
    long sheetIndex = definedName.getLocalSheetId();
    if (sheetIndex > Integer.MAX_VALUE
        || sheetIndex >= workbook.xssfWorkbook().getNumberOfSheets()) {
      return new NamedRangeScopeReport.Workbook();
    }
    return new NamedRangeScopeReport.Sheet(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(
            workbook.context().workbook().getSheetName((int) sheetIndex)));
  }

  private record NamedRangeIdentity(String name, NamedRangeScopeReport scope) {}
}
