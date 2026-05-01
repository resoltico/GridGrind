package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Name;

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
    List<ExcelNamedRangeSnapshot> namedRanges = new ArrayList<>();
    for (Name name : workbook.context().workbook().getAllNames()) {
      if (!shouldExpose(name)) {
        continue;
      }
      ExcelNamedRangeScope scope = toScope(workbook, name);
      String refersToFormula = Objects.requireNonNullElse(name.getRefersToFormula(), "");
      var target = ExcelNamedRangeTargets.resolveTarget(refersToFormula, scope);
      if (target.isEmpty()) {
        namedRanges.add(
            new ExcelNamedRangeSnapshot.FormulaSnapshot(
                name.getNameName(), scope, refersToFormula));
      } else {
        namedRanges.add(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                name.getNameName(), scope, refersToFormula, target.orElseThrow()));
      }
    }
    return List.copyOf(namedRanges);
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

  private static Name existingName(
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
}
