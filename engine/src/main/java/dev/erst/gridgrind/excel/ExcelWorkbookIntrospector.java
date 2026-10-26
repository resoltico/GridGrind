package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.formula.FormulaPatternReport;
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
import dev.erst.gridgrind.core.plan.InspectionIntrospectionResult;
import dev.erst.gridgrind.core.plan.InspectionSurfaceResult;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.result.TypeCountReport;
import dev.erst.gridgrind.core.vocabulary.NamedRangeBackingKind;
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
final class ExcelWorkbookIntrospector {
  private final ExcelSheetIntrospector sheetIntrospector;
  private final ExcelDocumentIntrospector documentIntrospector;

  ExcelWorkbookIntrospector() {
    this(new ExcelSheetIntrospector(), new ExcelDocumentIntrospector());
  }

  ExcelWorkbookIntrospector(
      ExcelSheetIntrospector sheetIntrospector, ExcelDocumentIntrospector documentIntrospector) {
    this.sheetIntrospector =
        Objects.requireNonNull(sheetIntrospector, "sheetIntrospector must not be null");
    this.documentIntrospector =
        Objects.requireNonNull(documentIntrospector, "documentIntrospector must not be null");
  }

  /** Executes one fact-only read command against the workbook. */
  InspectionIntrospectionResult execute(
      ExcelWorkbook workbook, WorkbookReadCommand.Introspection command) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(command, "command must not be null");

    return switch (command) {
      case WorkbookReadCommand.GetWorkbookSummary getWorkbookSummary ->
          new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult(
              getWorkbookSummary.stepId(), workbook.workbookSummary());
      case WorkbookReadCommand.GetPackageSecurity getPackageSecurity ->
          new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult(
              getPackageSecurity.stepId(), workbook.packageSecurity());
      case WorkbookReadCommand.GetWorkbookProtection getWorkbookProtection ->
          new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookProtectionResult(
              getWorkbookProtection.stepId(), workbook.workbookProtection());
      case WorkbookReadCommand.GetCustomXmlMappings getCustomXmlMappings ->
          new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.CustomXmlMappingsResult(
              getCustomXmlMappings.stepId(), workbook.customXml().customXmlMappings());
      case WorkbookReadCommand.ExportCustomXmlMapping exportCustomXmlMapping ->
          new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.CustomXmlExportResult(
              exportCustomXmlMapping.stepId(),
              workbook
                  .customXml()
                  .exportCustomXmlMapping(
                      exportCustomXmlMapping.mapping(),
                      exportCustomXmlMapping.validateSchema(),
                      exportCustomXmlMapping.encoding()));
      case WorkbookReadCommand.GetNamedRanges getNamedRanges ->
          new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.NamedRangesResult(
              getNamedRanges.stepId(), selectNamedRanges(workbook, getNamedRanges.selection()));
      case WorkbookReadCommand.GetSheetSummary getSheetSummary ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetSummaryResult(
              getSheetSummary.stepId(), workbook.sheetSummary(getSheetSummary.sheetName()));
      case WorkbookReadCommand.GetArrayFormulas getArrayFormulas ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.ArrayFormulasResult(
              getArrayFormulas.stepId(), arrayFormulas(workbook, getArrayFormulas.selection()));
      case WorkbookReadCommand.GetCells getCells ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.CellsResult(
              getCells.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(getCells.sheetName()),
              sheetIntrospector.cells(
                  workbook.sheet(getCells.sheetName()),
                  getCells.addresses(),
                  getCells.projection()));
      case WorkbookReadCommand.GetWindow getWindow ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.WindowResult(
              getWindow.stepId(), readWindow(workbook, getWindow));
      case WorkbookReadCommand.GetMergedRegions getMergedRegions ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.MergedRegionsResult(
              getMergedRegions.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(getMergedRegions.sheetName()),
              sheetIntrospector.mergedRegions(workbook.sheet(getMergedRegions.sheetName())));
      case WorkbookReadCommand.GetHyperlinks getHyperlinks ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.HyperlinksResult(
              getHyperlinks.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(getHyperlinks.sheetName()),
              sheetIntrospector.hyperlinks(
                  workbook.sheet(getHyperlinks.sheetName()), getHyperlinks.selection()));
      case WorkbookReadCommand.GetComments getComments ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.CommentsResult(
              getComments.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(getComments.sheetName()),
              sheetIntrospector.comments(
                  workbook.sheet(getComments.sheetName()), getComments.selection()));
      case WorkbookReadCommand.GetDrawingObjects getDrawingObjects ->
          new dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.DrawingObjectsResult(
              getDrawingObjects.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(getDrawingObjects.sheetName()),
              documentIntrospector.drawingObjects(workbook, getDrawingObjects.sheetName()));
      case WorkbookReadCommand.GetCharts getCharts ->
          new dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.ChartsResult(
              getCharts.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(
                  getCharts.selection().sheetName().value()),
              documentIntrospector.charts(workbook, getCharts.selection()));
      case WorkbookReadCommand.GetPivotTables getPivotTables ->
          new dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.PivotTablesResult(
              getPivotTables.stepId(),
              documentIntrospector.pivotTables(workbook, getPivotTables.selection()));
      case WorkbookReadCommand.GetDrawingObjectPayload getDrawingObjectPayload ->
          new dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.DrawingObjectPayloadResult(
              getDrawingObjectPayload.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(
                  getDrawingObjectPayload.sheetName()),
              documentIntrospector.drawingObjectPayload(
                  workbook,
                  getDrawingObjectPayload.sheetName(),
                  getDrawingObjectPayload.objectName()));
      case WorkbookReadCommand.GetSheetLayout getSheetLayout ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetLayoutResult(
              getSheetLayout.stepId(),
              sheetIntrospector.layout(workbook.sheet(getSheetLayout.sheetName())));
      case WorkbookReadCommand.GetPrintLayout getPrintLayout ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.PrintLayoutResult(
              getPrintLayout.stepId(),
              sheetIntrospector.printLayout(workbook.sheet(getPrintLayout.sheetName())));
      case WorkbookReadCommand.GetDataValidations getDataValidations ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.DataValidationsResult(
              getDataValidations.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(getDataValidations.sheetName()),
              documentIntrospector.dataValidations(
                  workbook.sheet(getDataValidations.sheetName()), getDataValidations.selection()));
      case WorkbookReadCommand.GetConditionalFormatting getConditionalFormatting ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.ConditionalFormattingResult(
              getConditionalFormatting.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(
                  getConditionalFormatting.sheetName()),
              documentIntrospector.conditionalFormatting(
                  workbook.sheet(getConditionalFormatting.sheetName()),
                  getConditionalFormatting.selection()));
      case WorkbookReadCommand.GetAutofilters getAutofilters ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.AutofiltersResult(
              getAutofilters.stepId(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(getAutofilters.sheetName()),
              documentIntrospector.autofilters(workbook, getAutofilters.sheetName()));
      case WorkbookReadCommand.GetTables getTables ->
          new dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.TablesResult(
              getTables.stepId(), documentIntrospector.tables(workbook, getTables.selection()));
    };
  }

  private dev.erst.gridgrind.core.model.layout.WindowReport readWindow(
      ExcelWorkbook workbook, WorkbookReadCommand.GetWindow query) {
    var dense =
        sheetIntrospector.window(
            workbook.sheet(query.sheetName()),
            query.window().topLeftAddress(),
            query.window().rowCount(),
            query.window().columnCount(),
            query.projection());
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

  /** Executes one derived factual surface-summary read command against the workbook. */
  InspectionSurfaceResult execute(ExcelWorkbook workbook, WorkbookReadCommand.Surface command) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(command, "command must not be null");

    return switch (command) {
      case WorkbookReadCommand.GetFormulaSurface getFormulaSurface ->
          new dev.erst.gridgrind.core.plan.WorkbookSurfaceInspectionResult.FormulaSurfaceResult(
              getFormulaSurface.stepId(), formulaSurface(workbook, getFormulaSurface.selection()));
      case WorkbookReadCommand.GetSheetSchema getSheetSchema ->
          new dev.erst.gridgrind.core.plan.WorkbookSurfaceInspectionResult.SheetSchemaResult(
              getSheetSchema.stepId(),
              sheetSchema(
                  workbook,
                  getSheetSchema.sheetName(),
                  getSheetSchema.window().topLeftAddress(),
                  getSheetSchema.window().rowCount(),
                  getSheetSchema.window().columnCount(),
                  getSheetSchema.projection(),
                  workbook.xssfWorkbook().isDate1904()));
      case WorkbookReadCommand.GetNamedRangeSurface getNamedRangeSurface ->
          new dev.erst.gridgrind.core.plan.WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult(
              getNamedRangeSurface.stepId(),
              namedRangeSurface(workbook, getNamedRangeSurface.selection()));
    };
  }

  /** Returns the named ranges selected by the provided workbook-core named-range selection. */
  List<NamedRangeReport> selectNamedRanges(
      ExcelWorkbook workbook, ExcelNamedRangeSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<NamedRangeReport> namedRanges = workbook.names().namedRanges();
    return switch (selection) {
      case ExcelNamedRangeSelection.All _ -> namedRanges;
      case ExcelNamedRangeSelection.Selected selected ->
          selectNamedRanges(namedRanges, selected.selectors());
    };
  }

  private List<NamedRangeReport> selectNamedRanges(
      List<NamedRangeReport> namedRanges, List<ExcelNamedRangeSelector> selectors) {
    Set<NamedRangeReport> matched = new LinkedHashSet<>();
    for (ExcelNamedRangeSelector selector : selectors) {
      matched.addAll(matchSelector(namedRanges, selector));
    }
    return List.copyOf(matched);
  }

  /** Returns the named ranges matched by one exact selector or throws when none match. */
  static List<NamedRangeReport> matchSelector(
      List<NamedRangeReport> namedRanges, ExcelNamedRangeSelector selector) {
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
      NamedRangeReport namedRange, ExcelNamedRangeSelector selector) {
    return switch (selector) {
      case ExcelNamedRangeSelector.ByName byName ->
          namedRange.name().value().equalsIgnoreCase(byName.name());
      case ExcelNamedRangeSelector.WorkbookScope workbookScope ->
          namedRange.name().value().equalsIgnoreCase(workbookScope.name())
              && namedRange.scope() instanceof NamedRangeScopeReport.Workbook;
      case ExcelNamedRangeSelector.SheetScope sheetScope ->
          namedRange.name().value().equalsIgnoreCase(sheetScope.name())
              && matchesSheetScope(namedRange.scope(), sheetScope.sheetName());
    };
  }

  private static boolean matchesSheetScope(NamedRangeScopeReport scope, String sheetName) {
    return scope instanceof NamedRangeScopeReport.Sheet namedRangeScope
        && namedRangeScope.sheetName().value().equals(sheetName);
  }

  private static NamedRangeNotFoundException notFound(ExcelNamedRangeSelector selector) {
    return switch (selector) {
      case ExcelNamedRangeSelector.ByName byName ->
          new NamedRangeNotFoundException(byName.name(), new NamedRangeScopeReport.Workbook());
      case ExcelNamedRangeSelector.WorkbookScope workbookScope ->
          new NamedRangeNotFoundException(
              workbookScope.name(), new NamedRangeScopeReport.Workbook());
      case ExcelNamedRangeSelector.SheetScope sheetScope ->
          new NamedRangeNotFoundException(
              sheetScope.name(),
              new NamedRangeScopeReport.Sheet(
                  new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetScope.sheetName())));
    };
  }

  /** Groups formulas by exact expression across the selected sheets. */
  @SuppressWarnings({"PMD.AvoidInstantiatingObjectsInLoops", "PMD.UseConcurrentHashMap"})
  FormulaSurfaceReport formulaSurface(ExcelWorkbook workbook, SheetSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<String> sheetNames = selectSheets(workbook, selection);
    List<SheetFormulaSurfaceReport> sheets = new ArrayList<>(sheetNames.size());
    int totalFormulaCellCount = 0;
    for (String sheetName : sheetNames) {
      ExcelSheet sheet = workbook.sheet(sheetName);
      List<ExcelFormulaCell> formulaCells = ExcelFormulaFactReader.formulaCells(sheet.xssfSheet());
      Map<String, List<String>> formulas = new LinkedHashMap<>();
      for (ExcelFormulaCell formulaCell : formulaCells) {
        formulas
            .computeIfAbsent(formulaCell.formula(), _ -> new ArrayList<>())
            .add(formulaCell.address());
      }
      List<FormulaPatternReport> patterns = new ArrayList<>(formulas.size());
      for (Map.Entry<String, List<String>> entry : formulas.entrySet()) {
        patterns.add(
            new FormulaPatternReport(
                entry.getKey(),
                entry.getValue().size(),
                List.copyOf(entry.getValue()).stream()
                    .map(dev.erst.gridgrind.core.vocabulary.CellAddress::new)
                    .toList()));
      }
      totalFormulaCellCount += formulaCells.size();
      sheets.add(
          new SheetFormulaSurfaceReport(
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName),
              formulaCells.size(),
              patterns.size(),
              List.copyOf(patterns)));
    }
    return new FormulaSurfaceReport(totalFormulaCellCount, List.copyOf(sheets));
  }

  private List<ArrayFormulaReport> arrayFormulas(ExcelWorkbook workbook, SheetSelector selection) {
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
      ExcelWorkbook workbook,
      String sheetName,
      String topLeftAddress,
      int rowCount,
      int columnCount,
      CellReadProjection projection,
      boolean date1904) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    ExcelSheet sheet = workbook.sheet(sheetName);
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
      ExcelWorkbook workbook, ExcelNamedRangeSelection selection) {
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

  private List<String> selectSheets(ExcelWorkbook workbook, SheetSelector selection) {
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
}
