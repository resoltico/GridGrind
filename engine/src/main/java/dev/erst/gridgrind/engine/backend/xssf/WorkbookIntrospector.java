package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.formula.FormulaSurfaceReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeSurfaceEntryReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeSurfaceReport;
import dev.erst.gridgrind.core.model.formula.SheetFormulaSurfaceReport;
import dev.erst.gridgrind.core.model.table.SchemaColumnReport;
import dev.erst.gridgrind.core.model.table.SheetSchemaReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.plan.CellReadFacet;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.result.TypeCountReport;
import dev.erst.gridgrind.core.vocabulary.NamedRangeBackingKind;
import dev.erst.gridgrind.engine.formulas.FormulaAdapter;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import org.apache.poi.ss.util.CellReference;
import org.jspecify.annotations.Nullable;

/** Reads workbook facts and sheet introspection data from one workbook wrapper. */
final class WorkbookIntrospector {
  private final SheetIntrospector sheetIntrospector;
  private final DocumentIntrospector documentIntrospector;

  WorkbookIntrospector() {
    this(new SheetIntrospector(), new DocumentIntrospector());
  }

  WorkbookIntrospector(
      SheetIntrospector sheetIntrospector, DocumentIntrospector documentIntrospector) {
    this.sheetIntrospector =
        Objects.requireNonNull(sheetIntrospector, "sheetIntrospector must not be null");
    this.documentIntrospector =
        Objects.requireNonNull(documentIntrospector, "documentIntrospector must not be null");
  }

  /** Executes one fact-only read command against the workbook. */
  dev.erst.gridgrind.core.plan.InspectionResult execute(
      XssfWorkbook workbook, dev.erst.gridgrind.core.plan.ResolvedInspectionStep command) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(command, "command must not be null");

    return switch (command.query()) {
      case dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.GetWorkbookSummary _ ->
          new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult(
              command.stepId(), workbook.workbookSummary());
      case dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.GetPackageSecurity _ ->
          new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult(
              command.stepId(), workbook.packageSecurity());
      case dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.GetWorkbookProtection _ ->
          new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookProtectionResult(
              command.stepId(), workbook.protection().snapshotWorkbook());
      case dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.GetCustomXmlMappings _ ->
          new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.CustomXmlMappingsResult(
              command.stepId(), workbook.customXml().mappings());
      case dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.ExportCustomXmlMapping
              exportCustomXmlMapping ->
          new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.CustomXmlExportResult(
              command.stepId(),
              workbook
                  .customXml()
                  .exportMapping(
                      exportCustomXmlMapping.mapping(),
                      exportCustomXmlMapping.validateSchema(),
                      exportCustomXmlMapping.encoding()));
      case dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.GetNamedRanges _ ->
          new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.NamedRangesResult(
              command.stepId(),
              selectNamedRanges(
                  workbook, (dev.erst.gridgrind.core.plan.NamedRangeSelector) command.target()));
      case dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetSheetSummary _ ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetSummaryResult(
              command.stepId(),
              workbook.sheetSummary(
                  ((dev.erst.gridgrind.core.plan.SheetSelector.ByName) command.target())
                      .name()
                      .value()));
      case dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetArrayFormulas _ ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.ArrayFormulasResult(
              command.stepId(),
              arrayFormulas(
                  workbook, (dev.erst.gridgrind.core.plan.SheetSelector) command.target()));
      case dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetCells getCells ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.CellsResult(
              command.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(cellSheet(command.target())),
              sheetIntrospector.cells(
                  workbook.sheet(cellSheet(command.target())),
                  cellAddresses(command.target()),
                  getCells.resolvedProjection()));
      case dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetWindow _ ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.WindowResult(
              command.stepId(), readWindow(workbook, command));
      case dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetMergedRegions _ ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.MergedRegionsResult(
              command.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(
                  ((dev.erst.gridgrind.core.plan.SheetSelector.ByName) command.target())
                      .name()
                      .value()),
              sheetIntrospector.mergedRegions(
                  workbook.sheet(
                      ((dev.erst.gridgrind.core.plan.SheetSelector.ByName) command.target())
                          .name()
                          .value())));
      case dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetHyperlinks _ ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.HyperlinksResult(
              command.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(cellSheet(command.target())),
              readHyperlinks(workbook, command.target()));
      case dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetComments _ ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.CommentsResult(
              command.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(cellSheet(command.target())),
              readComments(workbook, command.target()));
      case dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery.GetDrawingObjects _ ->
          new dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.DrawingObjectsResult(
              command.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(
                  ((dev.erst.gridgrind.core.plan.DrawingObjectSelector.AllOnSheet) command.target())
                      .sheetName()
                      .value()),
              documentIntrospector.drawingObjects(
                  workbook,
                  ((dev.erst.gridgrind.core.plan.DrawingObjectSelector.AllOnSheet) command.target())
                      .sheetName()
                      .value()));
      case dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery.GetCharts _ ->
          new dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.ChartsResult(
              command.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(
                  ((dev.erst.gridgrind.core.plan.ChartSelector) command.target())
                      .sheetName()
                      .value()),
              documentIntrospector.charts(
                  workbook, (dev.erst.gridgrind.core.plan.ChartSelector) command.target()));
      case dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery.GetPivotTables _ ->
          new dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.PivotTablesResult(
              command.stepId(),
              documentIntrospector.pivotTables(
                  workbook, (dev.erst.gridgrind.core.plan.PivotTableSelector) command.target()));
      case dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload _ ->
          new dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.DrawingObjectPayloadResult(
              command.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(
                  ((dev.erst.gridgrind.core.plan.DrawingObjectSelector.ByName) command.target())
                      .sheetName()
                      .value()),
              documentIntrospector.drawingObjectPayload(
                  workbook,
                  ((dev.erst.gridgrind.core.plan.DrawingObjectSelector.ByName) command.target())
                      .sheetName()
                      .value(),
                  ((dev.erst.gridgrind.core.plan.DrawingObjectSelector.ByName) command.target())
                      .objectName()));
      case dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetSheetLayout _ ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetLayoutResult(
              command.stepId(),
              sheetIntrospector.layout(
                  workbook.sheet(
                      ((dev.erst.gridgrind.core.plan.SheetSelector.ByName) command.target())
                          .name()
                          .value())));
      case dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetPrintLayout _ ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.PrintLayoutResult(
              command.stepId(),
              sheetIntrospector.printLayout(
                  workbook.sheet(
                      ((dev.erst.gridgrind.core.plan.SheetSelector.ByName) command.target())
                          .name()
                          .value())));
      case dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetDataValidations _ ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.DataValidationsResult(
              command.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(rangeSheet(command.target())),
              readDataValidations(workbook, command.target()));
      case dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetConditionalFormatting _ ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.ConditionalFormattingResult(
              command.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(rangeSheet(command.target())),
              readConditionalFormatting(workbook, command.target()));
      case dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetAutofilters _ ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.AutofiltersResult(
              command.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(
                  ((dev.erst.gridgrind.core.plan.SheetSelector.ByName) command.target())
                      .name()
                      .value()),
              documentIntrospector.autofilters(
                  workbook,
                  ((dev.erst.gridgrind.core.plan.SheetSelector.ByName) command.target())
                      .name()
                      .value()));
      case dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery.GetTables _ ->
          new dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.TablesResult(
              command.stepId(),
              documentIntrospector.tables(
                  workbook, (dev.erst.gridgrind.core.plan.TableSelector) command.target()));
      case dev.erst.gridgrind.core.plan.InspectionSurfaceQuery.GetFormulaSurface _ ->
          new dev.erst.gridgrind.core.plan.WorkbookSurfaceInspectionResult.FormulaSurfaceResult(
              command.stepId(),
              formulaSurface(
                  workbook, (dev.erst.gridgrind.core.plan.SheetSelector) command.target()));
      case dev.erst.gridgrind.core.plan.InspectionSurfaceQuery.GetSheetSchema getSheetSchema ->
          new dev.erst.gridgrind.core.plan.WorkbookSurfaceInspectionResult.SheetSchemaResult(
              command.stepId(),
              sheetSchema(
                  workbook,
                  ((dev.erst.gridgrind.core.plan.RangeSelector.RectangularWindow) command.target())
                      .sheetName()
                      .value(),
                  ((dev.erst.gridgrind.core.plan.RangeSelector.RectangularWindow) command.target())
                      .topLeftAddress()
                      .value(),
                  ((dev.erst.gridgrind.core.plan.RangeSelector.RectangularWindow) command.target())
                      .rowCount(),
                  ((dev.erst.gridgrind.core.plan.RangeSelector.RectangularWindow) command.target())
                      .columnCount(),
                  getSheetSchema.resolvedProjection(),
                  workbook.xssfWorkbook().isDate1904()));
      case dev.erst.gridgrind.core.plan.InspectionSurfaceQuery.GetNamedRangeSurface _ ->
          new dev.erst.gridgrind.core.plan.WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult(
              command.stepId(),
              namedRangeSurface(
                  workbook, (dev.erst.gridgrind.core.plan.NamedRangeSelector) command.target()));
      default ->
          throw new IllegalArgumentException(
              "Unsupported query group: " + command.query().queryType());
    };
  }

  private dev.erst.gridgrind.core.model.layout.WindowReport readWindow(
      XssfWorkbook workbook, dev.erst.gridgrind.core.plan.ResolvedInspectionStep command) {
    var query = (dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetWindow) command.query();
    var window = (dev.erst.gridgrind.core.plan.RangeSelector.RectangularWindow) command.target();
    var dense =
        sheetIntrospector.window(
            workbook.sheet(window.sheetName().value()),
            window.topLeftAddress().value(),
            window.rowCount(),
            window.columnCount(),
            query.resolvedProjection());
    if (query.includeBlanks()) return dense;
    return new dev.erst.gridgrind.core.model.layout.WindowReport.Sparse(
        dense.sheetName(),
        dense.topLeftAddress(),
        dense.dimensions(),
        dense.rows().stream()
            .flatMap(row -> row.cells().stream())
            .filter(
                cell ->
                    !(cell instanceof dev.erst.gridgrind.core.model.value.CellReport.BlankReport))
            .toList());
  }

  /** Returns the named ranges selected by the provided workbook-core named-range selection. */
  List<NamedRangeReport> selectNamedRanges(
      XssfWorkbook workbook, dev.erst.gridgrind.core.plan.NamedRangeSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<NamedRangeReport> namedRanges = workbook.names().namedRanges();
    return switch (selection) {
      case dev.erst.gridgrind.core.plan.NamedRangeSelector.All _ -> namedRanges;
      case dev.erst.gridgrind.core.plan.NamedRangeSelector.AnyOf selected ->
          selectNamedRanges(namedRanges, selected.selectors());
      case dev.erst.gridgrind.core.plan.NamedRangeSelector.ByNames selected ->
          selectNamedRanges(
              namedRanges,
              selected.names().stream()
                  .map(dev.erst.gridgrind.core.plan.NamedRangeSelector.ByName::new)
                  .toList());
      case dev.erst.gridgrind.core.plan.NamedRangeSelector.ByName selected ->
          matchSelector(namedRanges, selected);
      case dev.erst.gridgrind.core.plan.NamedRangeSelector.WorkbookScope selected ->
          matchSelector(namedRanges, selected);
      case dev.erst.gridgrind.core.plan.NamedRangeSelector.SheetScope selected ->
          matchSelector(namedRanges, selected);
    };
  }

  private List<NamedRangeReport> selectNamedRanges(
      List<NamedRangeReport> namedRanges,
      List<? extends dev.erst.gridgrind.core.plan.NamedRangeSelector.Ref> selectors) {
    Set<NamedRangeReport> matched = new LinkedHashSet<>();
    for (dev.erst.gridgrind.core.plan.NamedRangeSelector.Ref selector : selectors) {
      matched.addAll(matchSelector(namedRanges, selector));
    }
    return List.copyOf(matched);
  }

  /** Returns the named ranges matched by one exact selector or throws when none match. */
  static List<NamedRangeReport> matchSelector(
      List<NamedRangeReport> namedRanges,
      dev.erst.gridgrind.core.plan.NamedRangeSelector.Ref selector) {
    List<NamedRangeReport> matches = new ArrayList<>();
    for (NamedRangeReport namedRange : namedRanges) {
      if (matchesSelector(namedRange, selector)) {
        matches.add(namedRange);
      }
    }
    if (matches.isEmpty()) {
      throw notFound(selector);
    }
    return List.copyOf(matches);
  }

  private static boolean matchesSelector(
      NamedRangeReport namedRange, dev.erst.gridgrind.core.plan.NamedRangeSelector.Ref selector) {
    return switch (selector) {
      case dev.erst.gridgrind.core.plan.NamedRangeSelector.ByName byName ->
          namedRange.name().value().equalsIgnoreCase(byName.name().value());
      case dev.erst.gridgrind.core.plan.NamedRangeSelector.WorkbookScope workbookScope ->
          namedRange.name().value().equalsIgnoreCase(workbookScope.name().value())
              && namedRange.scope() instanceof NamedRangeScopeReport.Workbook;
      case dev.erst.gridgrind.core.plan.NamedRangeSelector.SheetScope sheetScope ->
          namedRange.name().value().equalsIgnoreCase(sheetScope.name().value())
              && matchesSheetScope(namedRange.scope(), sheetScope.sheetName().value());
    };
  }

  private static boolean matchesSheetScope(NamedRangeScopeReport scope, String sheetName) {
    return scope instanceof NamedRangeScopeReport.Sheet namedRangeScope
        && namedRangeScope.sheetName().value().equals(sheetName);
  }

  private static NamedRangeNotFoundException notFound(
      dev.erst.gridgrind.core.plan.NamedRangeSelector.Ref selector) {
    return switch (selector) {
      case dev.erst.gridgrind.core.plan.NamedRangeSelector.ByName byName ->
          new NamedRangeNotFoundException(
              byName.name().value(), new NamedRangeScopeReport.Workbook());
      case dev.erst.gridgrind.core.plan.NamedRangeSelector.WorkbookScope workbookScope ->
          new NamedRangeNotFoundException(
              workbookScope.name().value(), new NamedRangeScopeReport.Workbook());
      case dev.erst.gridgrind.core.plan.NamedRangeSelector.SheetScope sheetScope ->
          new NamedRangeNotFoundException(
              sheetScope.name().value(),
              new NamedRangeScopeReport.Sheet(
                  new dev.erst.gridgrind.core.vocabulary.SheetLabel(
                      sheetScope.sheetName().value())));
    };
  }

  /** Groups formulas by exact expression across the selected sheets. */
  FormulaSurfaceReport formulaSurface(XssfWorkbook workbook, SheetSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<String> sheetNames = selectSheets(workbook, selection);
    List<SheetFormulaSurfaceReport> sheets = new ArrayList<>(sheetNames.size());
    int totalFormulaCellCount = 0;
    for (String sheetName : sheetNames) {
      XssfSheet sheet = workbook.sheet(sheetName);
      SheetFormulaSurfaceReport surface = FormulaAdapter.surface(sheet.xssfSheet());
      totalFormulaCellCount += surface.formulaCellCount();
      sheets.add(surface);
    }
    return new FormulaSurfaceReport(totalFormulaCellCount, List.copyOf(sheets));
  }

  private List<ArrayFormulaReport> arrayFormulas(XssfWorkbook workbook, SheetSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<String> sheetNames = selectSheets(workbook, selection);
    List<ArrayFormulaReport> formulas = new ArrayList<>();
    for (String sheetName : sheetNames) {
      formulas.addAll(workbook.sheet(sheetName).cells().arrayFormulas());
    }
    return List.copyOf(formulas);
  }

  /** Infers a simple schema from the provided sheet window. */
  @SuppressWarnings({"PMD.AvoidInstantiatingObjectsInLoops", "PMD.UseConcurrentHashMap"})
  SheetSchemaReport sheetSchema(
      XssfWorkbook workbook,
      String sheetName,
      String topLeftAddress,
      int rowCount,
      int columnCount,
      CellReadProjection projection,
      boolean date1904) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    XssfSheet sheet = workbook.sheet(sheetName);
    dev.erst.gridgrind.core.model.layout.WindowReport.Dense window =
        sheetIntrospector.window(
            sheet,
            topLeftAddress,
            rowCount,
            columnCount,
            projectionIncludingFormulaEvaluation(projection));
    int topLeftColumn = new CellReference(topLeftAddress).getCol();
    boolean headerRowIsBlank =
        window.rows().getFirst().cells().stream()
            .allMatch(CellReport.BlankReport.class::isInstance);
    int dataRowCount = headerRowIsBlank ? 0 : Math.max(0, rowCount - 1);

    List<SchemaColumnReport> columns = new ArrayList<>(columnCount);
    for (int columnOffset = 0; columnOffset < columnCount; columnOffset++) {
      CellReport headerCell = window.rows().getFirst().cells().get(columnOffset);
      Map<String, Integer> typeCounts = new LinkedHashMap<>();
      int populatedCellCount = 0;
      int blankCellCount = 0;
      for (int rowIndex = 1; rowIndex < window.rows().size(); rowIndex++) {
        CellReport cell = window.rows().get(rowIndex).cells().get(columnOffset);
        switch (cell) {
          case CellReport.BlankReport _ -> blankCellCount++;
          case CellReport.FormulaReport f -> {
            populatedCellCount++;
            typeCounts.merge(
                schemaObservedValue(f.value().orElseThrow(), f.temporal(), projection),
                1,
                Integer::sum);
          }
          case CellReport.TextReport _ -> {
            populatedCellCount++;
            typeCounts.merge(schemaObservedType(cell, projection), 1, Integer::sum);
          }
          case CellReport.NumberReport _ -> {
            populatedCellCount++;
            typeCounts.merge(schemaObservedType(cell, projection), 1, Integer::sum);
          }
          case CellReport.BooleanReport _ -> {
            populatedCellCount++;
            typeCounts.merge(schemaObservedType(cell, projection), 1, Integer::sum);
          }
          case CellReport.ErrorReport _ -> {
            populatedCellCount++;
            typeCounts.merge(schemaObservedType(cell, projection), 1, Integer::sum);
          }
        }
      }

      List<TypeCountReport> observedTypes = new ArrayList<>(typeCounts.size());
      for (Map.Entry<String, Integer> entry : typeCounts.entrySet()) {
        observedTypes.add(new TypeCountReport(entry.getKey(), entry.getValue()));
      }

      columns.add(
          new SchemaColumnReport(
              topLeftColumn + columnOffset,
              new CellReference(
                      new CellReference(topLeftAddress).getRow(), topLeftColumn + columnOffset)
                  .formatAsString(),
              headerCell.displayValue().orElse(""),
              populatedCellCount,
              blankCellCount,
              List.copyOf(observedTypes),
              dominantType(typeCounts)));
    }

    return new SheetSchemaReport(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName),
        new dev.erst.gridgrind.core.vocabulary.CellAddress(topLeftAddress),
        rowCount,
        columnCount,
        dataRowCount,
        List.copyOf(columns));
  }

  /** Summarizes the scope and backing kind of selected named ranges. */
  NamedRangeSurfaceReport namedRangeSurface(
      XssfWorkbook workbook, dev.erst.gridgrind.core.plan.NamedRangeSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<NamedRangeReport> namedRanges = selectNamedRanges(workbook, selection);
    int workbookScopedCount = 0;
    int sheetScopedCount = 0;
    int rangeBackedCount = 0;
    int formulaBackedCount = 0;
    List<NamedRangeSurfaceEntryReport> entries = new ArrayList<>(namedRanges.size());
    for (NamedRangeReport namedRange : namedRanges) {
      NamedRangeBackingKind kind;
      switch (namedRange) {
        case NamedRangeReport.RangeReport _ -> {
          kind = NamedRangeBackingKind.RANGE;
          rangeBackedCount++;
        }
        case NamedRangeReport.FormulaReport _ -> {
          kind = NamedRangeBackingKind.FORMULA;
          formulaBackedCount++;
        }
      }
      switch (namedRange.scope()) {
        case NamedRangeScopeReport.Workbook _ -> workbookScopedCount++;
        case NamedRangeScopeReport.Sheet _ -> sheetScopedCount++;
      }
      entries.add(
          new NamedRangeSurfaceEntryReport(
              namedRange.name().value(), namedRange.scope(), namedRange.refersToFormula(), kind));
    }
    return new NamedRangeSurfaceReport(
        workbookScopedCount,
        sheetScopedCount,
        rangeBackedCount,
        formulaBackedCount,
        List.copyOf(entries));
  }

  private List<String> selectSheets(XssfWorkbook workbook, SheetSelector selection) {
    return switch (selection) {
      case SheetSelector.All _ -> workbook.sheets().sheetNames();
      case SheetSelector.ByNames selected ->
          selected.names().stream()
              .map(dev.erst.gridgrind.core.vocabulary.SheetName::value)
              .toList();
      case SheetSelector.ByName selected -> List.of(selected.name().value());
    };
  }

  private @Nullable String dominantType(Map<String, Integer> typeCounts) {
    String dominantType = null;
    int dominantCount = -1;
    boolean tie = false;
    for (Map.Entry<String, Integer> entry : typeCounts.entrySet()) {
      if (entry.getValue() > dominantCount) {
        dominantType = entry.getKey();
        dominantCount = entry.getValue();
        tie = false;
      } else if (entry.getValue() == dominantCount) {
        tie = true;
      }
    }
    return tie ? null : dominantType;
  }

  private static String schemaObservedType(CellReport snapshot, CellReadProjection projection) {
    return switch (snapshot) {
      case CellReport.TextReport _ -> "TEXT";
      case CellReport.NumberReport number ->
          projection.includes(CellReadFacet.TEMPORAL)
              ? temporalObservedType(number).orElse("NUMBER")
              : "NUMBER";
      case CellReport.BooleanReport _ -> "BOOLEAN";
      case CellReport.ErrorReport _ -> "ERROR";
      case CellReport.BlankReport _ -> "BLANK";
      case CellReport.FormulaReport _ ->
          throw new IllegalArgumentException(
              "Schema observed types must not receive FORMULA directly");
    };
  }

  private static String schemaObservedValue(
      dev.erst.gridgrind.core.model.value.CellValue value,
      Optional<dev.erst.gridgrind.core.model.value.CellTemporalReport> temporal,
      CellReadProjection projection) {
    return switch (value) {
      case dev.erst.gridgrind.core.model.value.CellValue.Blank _ -> "BLANK";
      case dev.erst.gridgrind.core.model.value.CellValue.Text _ -> "TEXT";
      case dev.erst.gridgrind.core.model.value.CellValue.BooleanValue _ -> "BOOLEAN";
      case dev.erst.gridgrind.core.model.value.CellValue.ErrorValue _ -> "ERROR";
      case dev.erst.gridgrind.core.model.value.CellValue.NumberValue _ ->
          projection.includes(CellReadFacet.TEMPORAL)
              ? temporal
                  .flatMap(dev.erst.gridgrind.core.model.value.CellTemporalReport::kind)
                  .map(Enum::name)
                  .orElse("NUMBER")
              : "NUMBER";
    };
  }

  private static Optional<String> temporalObservedType(CellReport.NumberReport snapshot) {
    return snapshot
        .temporal()
        .flatMap(dev.erst.gridgrind.core.model.value.CellTemporalReport::kind)
        .map(Enum::name);
  }

  private static CellReadProjection projectionIncludingFormulaEvaluation(
      CellReadProjection projection) {

    java.util.Set<CellReadFacet> facets = java.util.EnumSet.noneOf(CellReadFacet.class);
    for (CellReadFacet facet : CellReadFacet.values()) {
      if (projection.includes(facet)) {
        facets.add(facet);
      }
    }
    facets.add(CellReadFacet.VALUE);
    facets.add(CellReadFacet.FORMAT);
    return new CellReadProjection(List.copyOf(facets));
  }

  private static String cellSheet(dev.erst.gridgrind.core.plan.ResolvedTarget target) {
    return switch (target) {
      case dev.erst.gridgrind.core.plan.CellSelector.ByAddress cell -> cell.sheetName().value();
      case dev.erst.gridgrind.core.plan.CellSelector.ByAddresses cells -> cells.sheetName().value();
      case dev.erst.gridgrind.core.plan.CellSelector.AllUsedInSheet cells ->
          cells.sheetName().value();
      default -> throw new IllegalArgumentException("Expected resolved cell target");
    };
  }

  private static java.util.List<String> cellAddresses(
      dev.erst.gridgrind.core.plan.ResolvedTarget target) {
    return switch (target) {
      case dev.erst.gridgrind.core.plan.CellSelector.ByAddress cell ->
          java.util.List.of(cell.address().value());
      case dev.erst.gridgrind.core.plan.CellSelector.ByAddresses cells ->
          cells.addresses().stream()
              .map(dev.erst.gridgrind.core.vocabulary.CellAddress::value)
              .toList();
      default -> throw new IllegalArgumentException("Expected exact cell addresses");
    };
  }

  private static String rangeSheet(dev.erst.gridgrind.core.plan.ResolvedTarget target) {
    return switch (target) {
      case dev.erst.gridgrind.core.plan.RangeSelector.AllOnSheet range -> range.sheetName().value();
      case dev.erst.gridgrind.core.plan.RangeSelector.ByRange range -> range.sheetName().value();
      case dev.erst.gridgrind.core.plan.RangeSelector.ByRanges range -> range.sheetName().value();
      case dev.erst.gridgrind.core.plan.RangeSelector.RectangularWindow range ->
          range.sheetName().value();
      default -> throw new IllegalArgumentException("Expected resolved range target");
    };
  }

  private static List<dev.erst.gridgrind.core.model.value.CellHyperlinkReport> readHyperlinks(
      XssfWorkbook workbook, dev.erst.gridgrind.core.plan.ResolvedTarget target) {
    var annotations = workbook.sheet(cellSheet(target)).links();
    return switch (target) {
      case dev.erst.gridgrind.core.plan.CellSelector.AllUsedInSheet _ -> annotations.hyperlinks();
      case dev.erst.gridgrind.core.plan.CellSelector.ByAddress cell ->
          annotations.hyperlinks(List.of(cell.address()));
      case dev.erst.gridgrind.core.plan.CellSelector.ByAddresses cells ->
          annotations.hyperlinks(cells.addresses());
      default -> throw new IllegalArgumentException("Expected resolved cell selection");
    };
  }

  private static List<dev.erst.gridgrind.core.model.drawing.CellCommentReport> readComments(
      XssfWorkbook workbook, dev.erst.gridgrind.core.plan.ResolvedTarget target) {
    var annotations = workbook.sheet(cellSheet(target)).comments();
    return switch (target) {
      case dev.erst.gridgrind.core.plan.CellSelector.AllUsedInSheet _ -> annotations.comments();
      case dev.erst.gridgrind.core.plan.CellSelector.ByAddress cell ->
          annotations.comments(List.of(cell.address()));
      case dev.erst.gridgrind.core.plan.CellSelector.ByAddresses cells ->
          annotations.comments(cells.addresses());
      default -> throw new IllegalArgumentException("Expected resolved cell selection");
    };
  }

  private static List<dev.erst.gridgrind.core.model.validation.DataValidationEntryReport>
      readDataValidations(
          XssfWorkbook workbook, dev.erst.gridgrind.core.plan.ResolvedTarget target) {
    var metadata = workbook.sheet(rangeSheet(target)).metadata();
    return switch (target) {
      case dev.erst.gridgrind.core.plan.RangeSelector.AllOnSheet _ -> metadata.dataValidations();
      case dev.erst.gridgrind.core.plan.RangeSelector.ByRange range ->
          metadata.dataValidations(List.of(range.range()));
      case dev.erst.gridgrind.core.plan.RangeSelector.ByRanges ranges ->
          metadata.dataValidations(ranges.ranges());
      case dev.erst.gridgrind.core.plan.RangeSelector.RectangularWindow window ->
          metadata.dataValidations(List.of(window.range()));
      default -> throw new IllegalArgumentException("Expected resolved range target");
    };
  }

  private static List<dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport>
      readConditionalFormatting(
          XssfWorkbook workbook, dev.erst.gridgrind.core.plan.ResolvedTarget target) {
    var metadata = workbook.sheet(rangeSheet(target)).metadata();
    return switch (target) {
      case dev.erst.gridgrind.core.plan.RangeSelector.AllOnSheet _ ->
          metadata.conditionalFormatting();
      case dev.erst.gridgrind.core.plan.RangeSelector.ByRange range ->
          metadata.conditionalFormatting(List.of(range.range()));
      case dev.erst.gridgrind.core.plan.RangeSelector.ByRanges ranges ->
          metadata.conditionalFormatting(ranges.ranges());
      case dev.erst.gridgrind.core.plan.RangeSelector.RectangularWindow window ->
          metadata.conditionalFormatting(List.of(window.range()));
      default -> throw new IllegalArgumentException("Expected resolved range target");
    };
  }
}
