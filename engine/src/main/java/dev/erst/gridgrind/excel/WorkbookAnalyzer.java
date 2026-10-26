package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.formula.FormulaHealthReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeHealthReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableHealthReport;
import dev.erst.gridgrind.core.model.table.AutofilterHealthReport;
import dev.erst.gridgrind.core.model.table.TableHealthReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.core.model.validation.DataValidationHealthReport;
import dev.erst.gridgrind.core.model.value.HyperlinkHealthReport;
import dev.erst.gridgrind.core.plan.InspectionAnalysisResult;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.result.AnalysisSummaryReport;
import dev.erst.gridgrind.core.result.WorkbookFindingsReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.Pxg;
import org.apache.poi.ss.formula.ptg.Pxg3D;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;

/** Derives conclusion-bearing workbook findings from reusable workbook and sheet facts. */
@SuppressWarnings(
    "PMD.ExcessiveImports") // Closed analysis dispatcher shares the canonical report families.
final class WorkbookAnalyzer {
  private final ExcelWorkbookIntrospector workbookIntrospector;
  private final ExcelDocumentAnalyzer documentAnalyzer;

  WorkbookAnalyzer() {
    this(new ExcelWorkbookIntrospector(), new ExcelDocumentAnalyzer());
  }

  WorkbookAnalyzer(
      ExcelWorkbookIntrospector workbookIntrospector, ExcelDocumentAnalyzer documentAnalyzer) {
    this.workbookIntrospector =
        Objects.requireNonNull(workbookIntrospector, "workbookIntrospector must not be null");
    this.documentAnalyzer =
        Objects.requireNonNull(documentAnalyzer, "documentAnalyzer must not be null");
  }

  /** Executes one derived analysis command against the workbook. */
  InspectionAnalysisResult execute(
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      WorkbookReadCommand.Analysis command) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");
    Objects.requireNonNull(command, "command must not be null");

    return switch (command) {
      case WorkbookReadCommand.AnalyzeFormulaHealth analyzeFormulaHealth ->
          new dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.FormulaHealthResult(
              analyzeFormulaHealth.stepId(),
              formulaHealth(workbook, analyzeFormulaHealth.selection()));
      case WorkbookReadCommand.AnalyzeDataValidationHealth analyzeDataValidationHealth ->
          new dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.DataValidationHealthResult(
              analyzeDataValidationHealth.stepId(),
              dataValidationHealth(workbook, analyzeDataValidationHealth.selection()));
      case WorkbookReadCommand.AnalyzeConditionalFormattingHealth
              analyzeConditionalFormattingHealth ->
          new dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.ConditionalFormattingHealthResult(
              analyzeConditionalFormattingHealth.stepId(),
              conditionalFormattingHealth(
                  workbook, analyzeConditionalFormattingHealth.selection()));
      case WorkbookReadCommand.AnalyzeAutofilterHealth analyzeAutofilterHealth ->
          new dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.AutofilterHealthResult(
              analyzeAutofilterHealth.stepId(),
              autofilterHealth(workbook, analyzeAutofilterHealth.selection()));
      case WorkbookReadCommand.AnalyzeTableHealth analyzeTableHealth ->
          new dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.TableHealthResult(
              analyzeTableHealth.stepId(), tableHealth(workbook, analyzeTableHealth.selection()));
      case WorkbookReadCommand.AnalyzePivotTableHealth analyzePivotTableHealth ->
          new dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.PivotTableHealthResult(
              analyzePivotTableHealth.stepId(),
              pivotTableHealth(workbook, analyzePivotTableHealth.selection()));
      case WorkbookReadCommand.AnalyzeHyperlinkHealth analyzeHyperlinkHealth ->
          new dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.HyperlinkHealthResult(
              analyzeHyperlinkHealth.stepId(),
              hyperlinkHealth(workbook, workbookLocation, analyzeHyperlinkHealth.selection()));
      case WorkbookReadCommand.AnalyzeNamedRangeHealth analyzeNamedRangeHealth ->
          new dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.NamedRangeHealthResult(
              analyzeNamedRangeHealth.stepId(),
              namedRangeHealth(workbook, analyzeNamedRangeHealth.selection()));
      case WorkbookReadCommand.AnalyzeWorkbookFindings analyzeWorkbookFindings ->
          new dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.WorkbookFindingsResult(
              analyzeWorkbookFindings.stepId(), workbookFindings(workbook, workbookLocation));
    };
  }

  FormulaHealthReport formulaHealth(ExcelWorkbook workbook, SheetSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    int checkedFormulaCellCount = 0;
    List<AnalysisFindingReport> findings = new ArrayList<>();
    for (String sheetName : selectSheets(workbook, selection)) {
      ExcelSheet sheet = workbook.sheet(sheetName);
      checkedFormulaCellCount += sheet.diagnostics().formulaCellCount();
      findings.addAll(sheet.diagnostics().formulaHealthFindings());
    }
    return new FormulaHealthReport(
        checkedFormulaCellCount, summary(findings), List.copyOf(findings));
  }

  DataValidationHealthReport dataValidationHealth(ExcelWorkbook workbook, SheetSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return documentAnalyzer.dataValidationHealth(workbook, selection);
  }

  ConditionalFormattingHealthReport conditionalFormattingHealth(
      ExcelWorkbook workbook, SheetSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return documentAnalyzer.conditionalFormattingHealth(workbook, selection);
  }

  AutofilterHealthReport autofilterHealth(ExcelWorkbook workbook, SheetSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return documentAnalyzer.autofilterHealth(workbook, selection);
  }

  TableHealthReport tableHealth(ExcelWorkbook workbook, TableSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return documentAnalyzer.tableHealth(workbook, selection);
  }

  PivotTableHealthReport pivotTableHealth(ExcelWorkbook workbook, PivotTableSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return documentAnalyzer.pivotTableHealth(workbook, selection);
  }

  HyperlinkHealthReport hyperlinkHealth(
      ExcelWorkbook workbook, WorkbookLocation workbookLocation, SheetSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    int checkedHyperlinkCount = 0;
    List<AnalysisFindingReport> findings = new ArrayList<>();
    for (String sheetName : selectSheets(workbook, selection)) {
      ExcelSheet sheet = workbook.sheet(sheetName);
      checkedHyperlinkCount += sheet.diagnostics().hyperlinkCount();
      findings.addAll(sheet.diagnostics().hyperlinkHealthFindings(workbookLocation));
    }
    return new HyperlinkHealthReport(
        checkedHyperlinkCount, summary(findings), List.copyOf(findings));
  }

  /** Runs hyperlink-health analysis assuming the workbook has not yet been saved to disk. */
  HyperlinkHealthReport hyperlinkHealth(ExcelWorkbook workbook, SheetSelector selection) {
    return hyperlinkHealth(workbook, new WorkbookLocation.UnsavedWorkbook(), selection);
  }

  NamedRangeHealthReport namedRangeHealth(
      ExcelWorkbook workbook, ExcelNamedRangeSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<NamedRangeReport> namedRanges =
        workbookIntrospector.selectNamedRanges(workbook, selection);
    if (selection instanceof ExcelNamedRangeSelection.All
        && !workbook.wasMutatedSinceOpenInternal()) {
      List<NamedRangeReport> combined = new ArrayList<>(namedRanges);
      combined.addAll(ExcelWorkbookNamedRangeSupport.unmodeledObservedNamedRanges(workbook));
      namedRanges = List.copyOf(combined);
    }
    List<AnalysisFindingReport> findings = analyzeNamedRangeHealth(workbook, namedRanges);
    return new NamedRangeHealthReport(namedRanges.size(), summary(findings), List.copyOf(findings));
  }

  WorkbookFindingsReport workbookFindings(
      ExcelWorkbook workbook, WorkbookLocation workbookLocation) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");

    List<AnalysisFindingReport> combined = new ArrayList<>();
    combined.addAll(formulaHealth(workbook, new SheetSelector.All()).findings());
    combined.addAll(dataValidationHealth(workbook, new SheetSelector.All()).findings());
    combined.addAll(conditionalFormattingHealth(workbook, new SheetSelector.All()).findings());
    combined.addAll(autofilterHealth(workbook, new SheetSelector.All()).findings());
    combined.addAll(tableHealth(workbook, new TableSelector.All()).findings());
    combined.addAll(pivotTableHealth(workbook, new PivotTableSelector.All()).findings());
    combined.addAll(
        hyperlinkHealth(workbook, workbookLocation, new SheetSelector.All()).findings());
    combined.addAll(namedRangeHealth(workbook, new ExcelNamedRangeSelection.All()).findings());

    List<AnalysisFindingReport> findings =
        List.copyOf(new ArrayList<>(new LinkedHashSet<>(combined)));
    return new WorkbookFindingsReport(summary(findings), findings);
  }

  /** Aggregates workbook findings assuming the workbook has not yet been saved to disk. */
  WorkbookFindingsReport workbookFindings(ExcelWorkbook workbook) {
    return workbookFindings(workbook, new WorkbookLocation.UnsavedWorkbook());
  }

  private List<AnalysisFindingReport> analyzeNamedRangeHealth(
      ExcelWorkbook workbook, List<NamedRangeReport> namedRanges) {
    List<AnalysisFindingReport> findings = new ArrayList<>();
    Set<String> workbookSheetNames = new LinkedHashSet<>(workbook.sheets().sheetNames());
    for (NamedRangeReport namedRange : namedRanges) {
      AnalysisLocationReport.NamedRange location = namedRangeLocation(namedRange);
      String refersToFormula = namedRange.refersToFormula();
      String upperFormula = refersToFormula.toUpperCase(Locale.ROOT);
      if (upperFormula.contains("#REF!")) {
        findings.add(
            new AnalysisFindingReport(
                AnalysisFindingCode.NAMED_RANGE_BROKEN_REFERENCE,
                AnalysisSeverity.ERROR,
                "Named range contains a broken reference",
                "Named range formula contains #REF! and is broken.",
                location,
                List.of(refersToFormula)));
        continue;
      }

      switch (namedFormulaReferences(workbook, namedRange)) {
        case NamedFormulaReferences.Parsed parsed -> {
          for (String referencedSheet : parsed.sheetNames()) {
            if (!workbookSheetNames.contains(referencedSheet)) {
              findings.add(
                  new AnalysisFindingReport(
                      AnalysisFindingCode.NAMED_RANGE_BROKEN_REFERENCE,
                      AnalysisSeverity.ERROR,
                      "Named range targets a missing sheet",
                      "Named range refers to a sheet that does not exist: " + referencedSheet,
                      location,
                      List.of(refersToFormula)));
            }
          }
        }
        case NamedFormulaReferences.Unparseable _ ->
            findings.add(
                new AnalysisFindingReport(
                    AnalysisFindingCode.NAMED_RANGE_UNPARSEABLE_FORMULA,
                    AnalysisSeverity.ERROR,
                    "Named range formula could not be parsed",
                    "Named range formula could not be parsed with workbook context.",
                    location,
                    List.of(refersToFormula)));
      }
    }

    findings.addAll(scopeShadowingFindings(namedRanges));
    return List.copyOf(findings);
  }

  static List<AnalysisFindingReport> scopeShadowingFindings(List<NamedRangeReport> namedRanges) {
    List<AnalysisFindingReport> findings = new ArrayList<>();
    Set<String> workbookScopedNames = new LinkedHashSet<>();
    Set<String> sheetScopedNames = new LinkedHashSet<>();
    for (NamedRangeReport namedRange : namedRanges) {
      String normalized = namedRange.name().value().toUpperCase(Locale.ROOT);
      switch (namedRange.scope()) {
        case NamedRangeScopeReport.Workbook _ -> workbookScopedNames.add(normalized);
        case NamedRangeScopeReport.Sheet _ -> sheetScopedNames.add(normalized);
      }
    }

    for (NamedRangeReport namedRange : namedRanges) {
      String normalized = namedRange.name().value().toUpperCase(Locale.ROOT);
      if (workbookScopedNames.contains(normalized) && sheetScopedNames.contains(normalized)) {
        findings.add(
            new AnalysisFindingReport(
                AnalysisFindingCode.NAMED_RANGE_SCOPE_SHADOWING,
                AnalysisSeverity.INFO,
                "Named range scope shadowing",
                "This named range name exists in both workbook and sheet scope.",
                namedRangeLocation(namedRange),
                List.of(namedRange.refersToFormula())));
      }
    }
    return List.copyOf(new ArrayList<>(new LinkedHashSet<>(findings)));
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

  AnalysisSummaryReport summary(List<AnalysisFindingReport> findings) {
    return ExcelDocumentAnalyzer.summarizeFindings(findings);
  }

  private static AnalysisLocationReport.NamedRange namedRangeLocation(NamedRangeReport namedRange) {
    return new AnalysisLocationReport.NamedRange(namedRange.name().value(), namedRange.scope());
  }

  private NamedFormulaReferences namedFormulaReferences(
      ExcelWorkbook workbook, NamedRangeReport namedRange) {
    int sheetIndex = sheetIndex(workbook, namedRange.scope());
    try {
      Ptg[] tokens =
          FormulaParser.parse(
              namedRange.refersToFormula(),
              XSSFEvaluationWorkbook.create(workbook.xssfWorkbook()),
              FormulaType.CELL,
              sheetIndex,
              -1);
      Set<String> sheetNames = new LinkedHashSet<>();
      for (Ptg token : tokens) {
        if (token instanceof Pxg pxg && pxg.getExternalWorkbookNumber() < 1) {
          sheetNames.add(pxg.getSheetName());
          if (token instanceof Pxg3D pxg3D && pxg3D.getLastSheetName() != null) {
            sheetNames.add(pxg3D.getLastSheetName());
          }
        }
      }
      return new NamedFormulaReferences.Parsed(List.copyOf(sheetNames));
    } catch (RuntimeException exception) {
      return new NamedFormulaReferences.Unparseable();
    }
  }

  private static int sheetIndex(ExcelWorkbook workbook, NamedRangeScopeReport scope) {
    return switch (scope) {
      case NamedRangeScopeReport.Workbook _ -> 0;
      case NamedRangeScopeReport.Sheet sheetScope ->
          Math.max(0, workbook.xssfWorkbook().getSheetIndex(sheetScope.sheetName().value()));
    };
  }

  /** Parsed or unparseable sheet-reference facts for one formula-defined name. */
  private sealed interface NamedFormulaReferences
      permits NamedFormulaReferences.Parsed, NamedFormulaReferences.Unparseable {
    record Parsed(List<String> sheetNames) implements NamedFormulaReferences {}

    record Unparseable() implements NamedFormulaReferences {}
  }
}
