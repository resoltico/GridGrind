package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableSelection;
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
  WorkbookReadAnalysisResult execute(
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      WorkbookReadCommand.Analysis command) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");
    Objects.requireNonNull(command, "command must not be null");

    return switch (command) {
      case WorkbookReadCommand.AnalyzeFormulaHealth analyzeFormulaHealth ->
          new WorkbookAnalysisResult.FormulaHealthResult(
              analyzeFormulaHealth.stepId(),
              formulaHealth(workbook, analyzeFormulaHealth.selection()));
      case WorkbookReadCommand.AnalyzeDataValidationHealth analyzeDataValidationHealth ->
          new WorkbookAnalysisResult.DataValidationHealthResult(
              analyzeDataValidationHealth.stepId(),
              dataValidationHealth(workbook, analyzeDataValidationHealth.selection()));
      case WorkbookReadCommand.AnalyzeConditionalFormattingHealth
              analyzeConditionalFormattingHealth ->
          new WorkbookAnalysisResult.ConditionalFormattingHealthResult(
              analyzeConditionalFormattingHealth.stepId(),
              conditionalFormattingHealth(
                  workbook, analyzeConditionalFormattingHealth.selection()));
      case WorkbookReadCommand.AnalyzeAutofilterHealth analyzeAutofilterHealth ->
          new WorkbookAnalysisResult.AutofilterHealthResult(
              analyzeAutofilterHealth.stepId(),
              autofilterHealth(workbook, analyzeAutofilterHealth.selection()));
      case WorkbookReadCommand.AnalyzeTableHealth analyzeTableHealth ->
          new WorkbookAnalysisResult.TableHealthResult(
              analyzeTableHealth.stepId(), tableHealth(workbook, analyzeTableHealth.selection()));
      case WorkbookReadCommand.AnalyzePivotTableHealth analyzePivotTableHealth ->
          new WorkbookAnalysisResult.PivotTableHealthResult(
              analyzePivotTableHealth.stepId(),
              pivotTableHealth(workbook, analyzePivotTableHealth.selection()));
      case WorkbookReadCommand.AnalyzeHyperlinkHealth analyzeHyperlinkHealth ->
          new WorkbookAnalysisResult.HyperlinkHealthResult(
              analyzeHyperlinkHealth.stepId(),
              hyperlinkHealth(workbook, workbookLocation, analyzeHyperlinkHealth.selection()));
      case WorkbookReadCommand.AnalyzeNamedRangeHealth analyzeNamedRangeHealth ->
          new WorkbookAnalysisResult.NamedRangeHealthResult(
              analyzeNamedRangeHealth.stepId(),
              namedRangeHealth(workbook, analyzeNamedRangeHealth.selection()));
      case WorkbookReadCommand.AnalyzeWorkbookFindings analyzeWorkbookFindings ->
          new WorkbookAnalysisResult.WorkbookFindingsResult(
              analyzeWorkbookFindings.stepId(), workbookFindings(workbook, workbookLocation));
    };
  }

  WorkbookAnalysis.FormulaHealth formulaHealth(
      ExcelWorkbook workbook, ExcelSheetSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    int checkedFormulaCellCount = 0;
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (String sheetName : selectSheets(workbook, selection)) {
      ExcelSheet sheet = workbook.sheet(sheetName);
      checkedFormulaCellCount += sheet.diagnostics().formulaCellCount();
      findings.addAll(sheet.diagnostics().formulaHealthFindings());
    }
    return new WorkbookAnalysis.FormulaHealth(
        checkedFormulaCellCount, summary(findings), List.copyOf(findings));
  }

  WorkbookAnalysis.DataValidationHealth dataValidationHealth(
      ExcelWorkbook workbook, ExcelSheetSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return documentAnalyzer.dataValidationHealth(workbook, selection);
  }

  WorkbookAnalysis.ConditionalFormattingHealth conditionalFormattingHealth(
      ExcelWorkbook workbook, ExcelSheetSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return documentAnalyzer.conditionalFormattingHealth(workbook, selection);
  }

  WorkbookAnalysis.AutofilterHealth autofilterHealth(
      ExcelWorkbook workbook, ExcelSheetSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return documentAnalyzer.autofilterHealth(workbook, selection);
  }

  WorkbookAnalysis.TableHealth tableHealth(ExcelWorkbook workbook, ExcelTableSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return documentAnalyzer.tableHealth(workbook, selection);
  }

  WorkbookAnalysis.PivotTableHealth pivotTableHealth(
      ExcelWorkbook workbook, ExcelPivotTableSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return documentAnalyzer.pivotTableHealth(workbook, selection);
  }

  WorkbookAnalysis.HyperlinkHealth hyperlinkHealth(
      ExcelWorkbook workbook, WorkbookLocation workbookLocation, ExcelSheetSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    int checkedHyperlinkCount = 0;
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (String sheetName : selectSheets(workbook, selection)) {
      ExcelSheet sheet = workbook.sheet(sheetName);
      checkedHyperlinkCount += sheet.diagnostics().hyperlinkCount();
      findings.addAll(sheet.diagnostics().hyperlinkHealthFindings(workbookLocation));
    }
    return new WorkbookAnalysis.HyperlinkHealth(
        checkedHyperlinkCount, summary(findings), List.copyOf(findings));
  }

  /** Runs hyperlink-health analysis assuming the workbook has not yet been saved to disk. */
  WorkbookAnalysis.HyperlinkHealth hyperlinkHealth(
      ExcelWorkbook workbook, ExcelSheetSelection selection) {
    return hyperlinkHealth(workbook, new WorkbookLocation.UnsavedWorkbook(), selection);
  }

  WorkbookAnalysis.NamedRangeHealth namedRangeHealth(
      ExcelWorkbook workbook, ExcelNamedRangeSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<ExcelNamedRangeSnapshot> namedRanges =
        workbookIntrospector.selectNamedRanges(workbook, selection);
    if (selection instanceof ExcelNamedRangeSelection.All
        && !workbook.wasMutatedSinceOpenInternal()) {
      List<ExcelNamedRangeSnapshot> combined = new ArrayList<>(namedRanges);
      combined.addAll(ExcelWorkbookNamedRangeSupport.unmodeledObservedNamedRanges(workbook));
      namedRanges = List.copyOf(combined);
    }
    List<WorkbookAnalysis.AnalysisFinding> findings =
        analyzeNamedRangeHealth(workbook, namedRanges);
    return new WorkbookAnalysis.NamedRangeHealth(
        namedRanges.size(), summary(findings), List.copyOf(findings));
  }

  WorkbookAnalysis.WorkbookFindings workbookFindings(
      ExcelWorkbook workbook, WorkbookLocation workbookLocation) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");

    List<WorkbookAnalysis.AnalysisFinding> combined = new ArrayList<>();
    combined.addAll(formulaHealth(workbook, new ExcelSheetSelection.All()).findings());
    combined.addAll(dataValidationHealth(workbook, new ExcelSheetSelection.All()).findings());
    combined.addAll(
        conditionalFormattingHealth(workbook, new ExcelSheetSelection.All()).findings());
    combined.addAll(autofilterHealth(workbook, new ExcelSheetSelection.All()).findings());
    combined.addAll(tableHealth(workbook, new ExcelTableSelection.All()).findings());
    combined.addAll(pivotTableHealth(workbook, new ExcelPivotTableSelection.All()).findings());
    combined.addAll(
        hyperlinkHealth(workbook, workbookLocation, new ExcelSheetSelection.All()).findings());
    combined.addAll(namedRangeHealth(workbook, new ExcelNamedRangeSelection.All()).findings());

    List<WorkbookAnalysis.AnalysisFinding> findings =
        List.copyOf(new ArrayList<>(new LinkedHashSet<>(combined)));
    return new WorkbookAnalysis.WorkbookFindings(summary(findings), findings);
  }

  /** Aggregates workbook findings assuming the workbook has not yet been saved to disk. */
  WorkbookAnalysis.WorkbookFindings workbookFindings(ExcelWorkbook workbook) {
    return workbookFindings(workbook, new WorkbookLocation.UnsavedWorkbook());
  }

  private List<WorkbookAnalysis.AnalysisFinding> analyzeNamedRangeHealth(
      ExcelWorkbook workbook, List<ExcelNamedRangeSnapshot> namedRanges) {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    Set<String> workbookSheetNames = new LinkedHashSet<>(workbook.sheets().sheetNames());
    for (ExcelNamedRangeSnapshot namedRange : namedRanges) {
      WorkbookAnalysis.AnalysisLocation.NamedRange location = namedRangeLocation(namedRange);
      String refersToFormula = namedRange.refersToFormula();
      String upperFormula = refersToFormula.toUpperCase(Locale.ROOT);
      if (upperFormula.contains("#REF!")) {
        findings.add(
            new WorkbookAnalysis.AnalysisFinding(
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
                  new WorkbookAnalysis.AnalysisFinding(
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
                new WorkbookAnalysis.AnalysisFinding(
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

  static List<WorkbookAnalysis.AnalysisFinding> scopeShadowingFindings(
      List<ExcelNamedRangeSnapshot> namedRanges) {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    Set<String> workbookScopedNames = new LinkedHashSet<>();
    Set<String> sheetScopedNames = new LinkedHashSet<>();
    for (ExcelNamedRangeSnapshot namedRange : namedRanges) {
      String normalized = namedRange.name().toUpperCase(Locale.ROOT);
      switch (namedRange.scope()) {
        case ExcelNamedRangeScope.WorkbookScope _ -> workbookScopedNames.add(normalized);
        case ExcelNamedRangeScope.SheetScope _ -> sheetScopedNames.add(normalized);
      }
    }

    for (ExcelNamedRangeSnapshot namedRange : namedRanges) {
      String normalized = namedRange.name().toUpperCase(Locale.ROOT);
      if (workbookScopedNames.contains(normalized) && sheetScopedNames.contains(normalized)) {
        findings.add(
            new WorkbookAnalysis.AnalysisFinding(
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

  private List<String> selectSheets(ExcelWorkbook workbook, ExcelSheetSelection selection) {
    return switch (selection) {
      case ExcelSheetSelection.All _ -> workbook.sheets().sheetNames();
      case ExcelSheetSelection.Selected selected -> List.copyOf(selected.sheetNames());
    };
  }

  WorkbookAnalysis.AnalysisSummary summary(List<WorkbookAnalysis.AnalysisFinding> findings) {
    return ExcelDocumentAnalyzer.summarizeFindings(findings);
  }

  private static WorkbookAnalysis.AnalysisLocation.NamedRange namedRangeLocation(
      ExcelNamedRangeSnapshot namedRange) {
    return new WorkbookAnalysis.AnalysisLocation.NamedRange(namedRange.name(), namedRange.scope());
  }

  private NamedFormulaReferences namedFormulaReferences(
      ExcelWorkbook workbook, ExcelNamedRangeSnapshot namedRange) {
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

  private static int sheetIndex(ExcelWorkbook workbook, ExcelNamedRangeScope scope) {
    return switch (scope) {
      case ExcelNamedRangeScope.WorkbookScope _ -> 0;
      case ExcelNamedRangeScope.SheetScope sheetScope ->
          Math.max(0, workbook.xssfWorkbook().getSheetIndex(sheetScope.sheetName()));
    };
  }

  /** Parsed or unparseable sheet-reference facts for one formula-defined name. */
  private sealed interface NamedFormulaReferences
      permits NamedFormulaReferences.Parsed, NamedFormulaReferences.Unparseable {
    record Parsed(List<String> sheetNames) implements NamedFormulaReferences {}

    record Unparseable() implements NamedFormulaReferences {}
  }
}
