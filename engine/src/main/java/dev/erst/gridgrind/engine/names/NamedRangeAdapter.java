package dev.erst.gridgrind.engine.names;

import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.engine.backend.xssf.NamedRangeNotFoundException;
import dev.erst.gridgrind.engine.backend.xssf.SheetCollectionAdapter;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.poi.CellBounds;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedNames;

/** Owns defined-name authoring, observed readback, scopes and reference normalization. */
public final class NamedRangeAdapter {
  private final XssfWorkbook workbook;

  /** Binds defined-name operations to one opened workbook. */
  public NamedRangeAdapter(XssfWorkbook workbook) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
  }

  /** Returns the number of names exposed by the workbook model. */
  public int namedRangeCount() {
    return namedRanges().size();
  }

  /** Creates or replaces a name in its authoritative workbook or sheet scope. */
  public NamedRangeAdapter setNamedRange(StructuredMutationAction.SetNamedRange definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    int index =
        switch (definition.scope()) {
          case dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook _ -> -1;
          case dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet sheet ->
              SheetCollectionAdapter.requiredSheetIndex(
                  workbook.xssfWorkbook(), sheet.sheetName().value());
        };
    String formula = toFormula(definition.target());
    Name name = existingName(workbook, definition.name().value(), index);
    if (name == null) name = workbook.xssfWorkbook().createName();
    name.setSheetIndex(index);
    name.setNameName(definition.name().value());
    name.setRefersToFormula(formula);
    return this;
  }

  /** Removes the matching scoped name. */
  public NamedRangeAdapter deleteNamedRange(String name, NamedRangeScopeReport scope) {
    Name existingName = requiredName(workbook, name, scope);
    workbook.xssfWorkbook().removeName(existingName);
    return this;
  }

  /** Returns names exposed by the workbook model. */
  public List<NamedRangeReport> namedRanges() {
    return namedRangesFromPoi(workbook);
  }

  /** Returns raw existing names that Apache POI did not model as {@link Name} instances. */
  public List<NamedRangeReport> unmodeledObservedNamedRanges() {
    CTDefinedNames definedNames = workbook.xssfWorkbook().getCTWorkbook().getDefinedNames();
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

  private static List<NamedRangeReport> namedRangesFromPoi(XssfWorkbook workbook) {
    List<NamedRangeReport> namedRanges = new ArrayList<>();
    for (Name name : workbook.xssfWorkbook().getAllNames()) {
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
    var target = resolveTarget(refersToFormula, scope);
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

  /** Checks an observed name against one explicit scope. */
  public boolean scopeMatches(Name candidate, NamedRangeScopeReport scope) {
    return switch (scope) {
      case NamedRangeScopeReport.Workbook _ -> candidate.getSheetIndex() < 0;
      case NamedRangeScopeReport.Sheet sheetScope ->
          candidate.getSheetIndex()
              == SheetCollectionAdapter.requiredSheetIndex(
                  workbook.xssfWorkbook(), sheetScope.sheetName().value());
    };
  }

  /** Checks whether an observed POI name belongs to the factual surface. */
  public static boolean shouldExpose(Name name) {
    return shouldExpose(name.getNameName(), name.isFunctionName(), name.isHidden());
  }

  /** Excludes function, hidden and built-in workbook names. */
  public static boolean shouldExpose(String nameName, boolean functionName, boolean hidden) {
    return !functionName
        && !hidden
        && nameName != null
        && !nameName.startsWith("_xlnm.")
        && !nameName.startsWith("_XLNM.");
  }

  private static Name requiredName(
      XssfWorkbook workbook, String name, NamedRangeScopeReport scope) {
    Name existingName = existingName(workbook, name, scopeIndex(workbook, scope));
    if (existingName == null) {
      throw new NamedRangeNotFoundException(name, scope);
    }
    return existingName;
  }

  private static @Nullable Name existingName(XssfWorkbook workbook, String name, int scopeIndex) {
    String validatedName =
        dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName(name);
    return workbook.xssfWorkbook().getAllNames().stream()
        .filter(candidate -> candidate.getNameName().equalsIgnoreCase(validatedName))
        .filter(
            candidate ->
                scopeIndex < 0
                    ? candidate.getSheetIndex() < 0
                    : candidate.getSheetIndex() == scopeIndex)
        .findFirst()
        .orElse(null);
  }

  private static int scopeIndex(XssfWorkbook workbook, NamedRangeScopeReport scope) {
    return switch (scope) {
      case NamedRangeScopeReport.Workbook _ -> -1;
      case NamedRangeScopeReport.Sheet sheet ->
          SheetCollectionAdapter.requiredSheetIndex(
              workbook.xssfWorkbook(), sheet.sheetName().value());
    };
  }

  private static NamedRangeScopeReport toScope(XssfWorkbook workbook, Name name) {
    int sheetIndex = name.getSheetIndex();
    if (sheetIndex < 0) {
      return new NamedRangeScopeReport.Workbook();
    }
    return new NamedRangeScopeReport.Sheet(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(
            workbook.xssfWorkbook().getSheetName(sheetIndex)));
  }

  private static NamedRangeScopeReport toScope(XssfWorkbook workbook, CTDefinedName definedName) {
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
            workbook.xssfWorkbook().getSheetName((int) sheetIndex)));
  }

  private record NamedRangeIdentity(String name, NamedRangeScopeReport scope) {}

  /** Formats an authored range target into its stored defined-name expression. */
  public static String toFormula(NamedRangeTarget target) {
    return switch (target) {
      case NamedRangeTarget.Formula formula -> formula.formula();
      case NamedRangeTarget.Range range -> rangeFormula(range.sheetName(), range.range());
    };
  }

  /** Formats a concrete range for a destination sheet at the POI boundary. */
  public static String rangeFormula(
      dev.erst.gridgrind.core.vocabulary.SheetName sheetName,
      dev.erst.gridgrind.core.vocabulary.RangeRef range) {
    CellBounds parsed = CellBounds.parse(range.value());
    String first =
        new CellReference(sheetName.value(), parsed.firstRow(), parsed.firstColumn(), true, true)
            .formatAsString();
    return parsed.rowCount() == 1 && parsed.columnCount() == 1
        ? first
        : first
            + ":"
            + new CellReference(parsed.lastRow(), parsed.lastColumn(), true, true).formatAsString();
  }

  /** Resolves a single local rectangle, retaining nonrectangular formulas as opaque facts. */
  public static Optional<NamedRangeTarget> resolveTarget(
      String refersToFormula, NamedRangeScopeReport scope) {
    Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
    Objects.requireNonNull(scope, "scope must not be null");

    if (refersToFormula.isBlank()) {
      return Optional.empty();
    }
    String normalizedFormula = normalizeAreaFormulaForPoi(refersToFormula);
    if (!AreaReference.isContiguous(normalizedFormula)) {
      return Optional.empty();
    }
    try {
      AreaReference areaReference =
          new AreaReference(normalizedFormula, SpreadsheetVersion.EXCEL2007);
      String sheetName = resolveSheetName(areaReference, scope);
      if (sheetName == null) {
        return Optional.empty();
      }
      CellReference firstCell = areaReference.getFirstCell();
      CellReference lastCell = areaReference.getLastCell();
      String range =
          firstCell.getRow() == lastCell.getRow() && firstCell.getCol() == lastCell.getCol()
              ? new CellReference(firstCell.getRow(), firstCell.getCol()).formatAsString()
              : new CellReference(firstCell.getRow(), firstCell.getCol()).formatAsString()
                  + ":"
                  + new CellReference(lastCell.getRow(), lastCell.getCol()).formatAsString();
      return Optional.of(
          NamedRangeTarget.range(
              new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName),
              new dev.erst.gridgrind.core.vocabulary.RangeRef(range)));
    } catch (RuntimeException exception) {
      return Optional.empty();
    }
  }

  /** Removes only an identical repeated sheet prefix from a two-cell reference. */
  public static String normalizeAreaFormulaForPoi(String refersToFormula) {
    int separatorIndex = refersToFormula.indexOf(':');
    if (separatorIndex < 0) {
      return refersToFormula;
    }
    String firstReference = refersToFormula.substring(0, separatorIndex);
    String secondReference = refersToFormula.substring(separatorIndex + 1);
    int firstBangIndex = firstReference.lastIndexOf('!');
    int secondBangIndex = secondReference.lastIndexOf('!');
    if (firstBangIndex < 0 || secondBangIndex < 0) {
      return refersToFormula;
    }
    String firstSheetPrefix = firstReference.substring(0, firstBangIndex + 1);
    String secondSheetPrefix = secondReference.substring(0, secondBangIndex + 1);
    if (!firstSheetPrefix.equals(secondSheetPrefix)) {
      return refersToFormula;
    }
    return firstReference + ":" + secondReference.substring(secondBangIndex + 1);
  }

  private static @Nullable String resolveSheetName(
      AreaReference areaReference, NamedRangeScopeReport scope) {
    String sheetName = areaReference.getFirstCell().getSheetName();
    if (sheetName != null) {
      return sheetName;
    }
    return switch (scope) {
      case NamedRangeScopeReport.Workbook _ -> null;
      case NamedRangeScopeReport.Sheet sheetScope -> sheetScope.sheetName().value();
    };
  }
}
