package dev.erst.gridgrind.excel;

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

  static ExcelWorkbook setNamedRange(ExcelWorkbook workbook, ExcelNamedRangeDefinition definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    Name name = existingName(workbook, definition.name(), definition.scope());
    if (name == null) {
      name = workbook.context().workbook().createName();
    }
    applyScope(workbook, name, definition.scope());
    name.setNameName(definition.name());
    name.setRefersToFormula(definition.target().refersToFormula());
    return workbook;
  }

  static ExcelWorkbook deleteNamedRange(
      ExcelWorkbook workbook, String name, ExcelNamedRangeScope scope) {
    Name existingName = requiredName(workbook, name, scope);
    workbook.context().workbook().removeName(existingName);
    return workbook;
  }

  static List<ExcelNamedRangeSnapshot> namedRanges(ExcelWorkbook workbook) {
    return namedRangesFromPoi(workbook);
  }

  /** Returns raw existing names that Apache POI did not model as {@link Name} instances. */
  static List<ExcelNamedRangeSnapshot> unmodeledObservedNamedRanges(ExcelWorkbook workbook) {
    CTDefinedNames definedNames = workbook.context().workbook().getCTWorkbook().getDefinedNames();
    if (definedNames == null) {
      return List.of();
    }
    Set<NamedRangeIdentity> modeledNames = new LinkedHashSet<>();
    for (ExcelNamedRangeSnapshot snapshot : namedRangesFromPoi(workbook)) {
      modeledNames.add(new NamedRangeIdentity(snapshot.name(), snapshot.scope()));
    }
    List<ExcelNamedRangeSnapshot> namedRanges = new ArrayList<>();
    for (CTDefinedName definedName : definedNames.getDefinedNameList()) {
      String name = Objects.requireNonNullElse(definedName.getName(), "");
      if (name.isBlank()
          || !shouldExpose(name, definedName.getFunction(), definedName.getHidden())) {
        continue;
      }
      ExcelNamedRangeScope scope = toScope(workbook, definedName);
      if (modeledNames.add(new NamedRangeIdentity(name, scope))) {
        namedRanges.add(
            snapshot(name, scope, Objects.requireNonNullElse(definedName.getStringValue(), "")));
      }
    }
    return List.copyOf(namedRanges);
  }

  private static List<ExcelNamedRangeSnapshot> namedRangesFromPoi(ExcelWorkbook workbook) {
    List<ExcelNamedRangeSnapshot> namedRanges = new ArrayList<>();
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

  private static ExcelNamedRangeSnapshot snapshot(
      String name, ExcelNamedRangeScope scope, String refersToFormula) {
    var target = ExcelNamedRangeTargets.resolveTarget(refersToFormula, scope);
    if (target.isEmpty()) {
      return new ExcelNamedRangeSnapshot.FormulaSnapshot(name, scope, refersToFormula);
    }
    return new ExcelNamedRangeSnapshot.RangeSnapshot(
        name, scope, refersToFormula, target.orElseThrow());
  }

  static boolean scopeMatches(ExcelWorkbook workbook, Name candidate, ExcelNamedRangeScope scope) {
    return switch (scope) {
      case ExcelNamedRangeScope.WorkbookScope _ -> candidate.getSheetIndex() < 0;
      case ExcelNamedRangeScope.SheetScope sheetScope ->
          candidate.getSheetIndex() == workbook.requiredSheetIndex(sheetScope.sheetName());
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
      ExcelWorkbook workbook, String name, ExcelNamedRangeScope scope) {
    Name existingName = existingName(workbook, name, scope);
    if (existingName == null) {
      throw new NamedRangeNotFoundException(name, scope);
    }
    return existingName;
  }

  private static @Nullable Name existingName(
      ExcelWorkbook workbook, String name, ExcelNamedRangeScope scope) {
    String validatedName = ExcelNamedRangeDefinition.validateName(name);
    Objects.requireNonNull(scope, "scope must not be null");

    return workbook.context().workbook().getAllNames().stream()
        .filter(candidate -> candidate.getNameName().equalsIgnoreCase(validatedName))
        .filter(candidate -> scopeMatches(workbook, candidate, scope))
        .findFirst()
        .orElse(null);
  }

  private static void applyScope(ExcelWorkbook workbook, Name name, ExcelNamedRangeScope scope) {
    switch (scope) {
      case ExcelNamedRangeScope.WorkbookScope _ -> name.setSheetIndex(-1);
      case ExcelNamedRangeScope.SheetScope sheetScope ->
          name.setSheetIndex(workbook.requiredSheetIndex(sheetScope.sheetName()));
    }
  }

  private static ExcelNamedRangeScope toScope(ExcelWorkbook workbook, Name name) {
    int sheetIndex = name.getSheetIndex();
    if (sheetIndex < 0) {
      return new ExcelNamedRangeScope.WorkbookScope();
    }
    return new ExcelNamedRangeScope.SheetScope(
        workbook.context().workbook().getSheetName(sheetIndex));
  }

  private static ExcelNamedRangeScope toScope(ExcelWorkbook workbook, CTDefinedName definedName) {
    if (!definedName.isSetLocalSheetId()) {
      return new ExcelNamedRangeScope.WorkbookScope();
    }
    long sheetIndex = definedName.getLocalSheetId();
    if (sheetIndex > Integer.MAX_VALUE
        || sheetIndex >= workbook.xssfWorkbook().getNumberOfSheets()) {
      return new ExcelNamedRangeScope.WorkbookScope();
    }
    return new ExcelNamedRangeScope.SheetScope(
        workbook.context().workbook().getSheetName((int) sheetIndex));
  }

  private record NamedRangeIdentity(String name, ExcelNamedRangeScope scope) {}
}
