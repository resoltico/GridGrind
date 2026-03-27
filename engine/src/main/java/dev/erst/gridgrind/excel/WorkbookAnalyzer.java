package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** Derives conclusion-bearing workbook findings from reusable workbook and sheet facts. */
final class WorkbookAnalyzer {
  private final ExcelWorkbookIntrospector workbookIntrospector;

  WorkbookAnalyzer() {
    this(new ExcelWorkbookIntrospector());
  }

  WorkbookAnalyzer(ExcelWorkbookIntrospector workbookIntrospector) {
    this.workbookIntrospector =
        Objects.requireNonNull(workbookIntrospector, "workbookIntrospector must not be null");
  }

  /** Executes one derived analysis command against the workbook. */
  WorkbookReadResult.Analysis execute(
      ExcelWorkbook workbook, WorkbookReadCommand.Analysis command) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(command, "command must not be null");

    return switch (command) {
      case WorkbookReadCommand.AnalyzeFormulaHealth analyzeFormulaHealth ->
          new WorkbookReadResult.FormulaHealthResult(
              analyzeFormulaHealth.requestId(),
              formulaHealth(workbook, analyzeFormulaHealth.selection()));
      case WorkbookReadCommand.AnalyzeHyperlinkHealth analyzeHyperlinkHealth ->
          new WorkbookReadResult.HyperlinkHealthResult(
              analyzeHyperlinkHealth.requestId(),
              hyperlinkHealth(workbook, analyzeHyperlinkHealth.selection()));
      case WorkbookReadCommand.AnalyzeNamedRangeHealth analyzeNamedRangeHealth ->
          new WorkbookReadResult.NamedRangeHealthResult(
              analyzeNamedRangeHealth.requestId(),
              namedRangeHealth(workbook, analyzeNamedRangeHealth.selection()));
      case WorkbookReadCommand.AnalyzeWorkbookFindings analyzeWorkbookFindings ->
          new WorkbookReadResult.WorkbookFindingsResult(
              analyzeWorkbookFindings.requestId(), workbookFindings(workbook));
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
      checkedFormulaCellCount += sheet.formulaCellCount();
      findings.addAll(sheet.formulaHealthFindings());
    }
    return new WorkbookAnalysis.FormulaHealth(
        checkedFormulaCellCount, summary(findings), List.copyOf(findings));
  }

  WorkbookAnalysis.HyperlinkHealth hyperlinkHealth(
      ExcelWorkbook workbook, ExcelSheetSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    int checkedHyperlinkCount = 0;
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (String sheetName : selectSheets(workbook, selection)) {
      ExcelSheet sheet = workbook.sheet(sheetName);
      checkedHyperlinkCount += sheet.hyperlinkCount();
      findings.addAll(sheet.hyperlinkHealthFindings());
    }
    return new WorkbookAnalysis.HyperlinkHealth(
        checkedHyperlinkCount, summary(findings), List.copyOf(findings));
  }

  WorkbookAnalysis.NamedRangeHealth namedRangeHealth(
      ExcelWorkbook workbook, ExcelNamedRangeSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<ExcelNamedRangeSnapshot> namedRanges =
        workbookIntrospector.selectNamedRanges(workbook, selection);
    List<WorkbookAnalysis.AnalysisFinding> findings =
        analyzeNamedRangeHealth(workbook, namedRanges);
    return new WorkbookAnalysis.NamedRangeHealth(
        namedRanges.size(), summary(findings), List.copyOf(findings));
  }

  WorkbookAnalysis.WorkbookFindings workbookFindings(ExcelWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");

    List<WorkbookAnalysis.AnalysisFinding> combined = new ArrayList<>();
    combined.addAll(formulaHealth(workbook, new ExcelSheetSelection.All()).findings());
    combined.addAll(hyperlinkHealth(workbook, new ExcelSheetSelection.All()).findings());
    combined.addAll(namedRangeHealth(workbook, new ExcelNamedRangeSelection.All()).findings());

    List<WorkbookAnalysis.AnalysisFinding> findings =
        List.copyOf(new ArrayList<>(new LinkedHashSet<>(combined)));
    return new WorkbookAnalysis.WorkbookFindings(summary(findings), findings);
  }

  private List<WorkbookAnalysis.AnalysisFinding> analyzeNamedRangeHealth(
      ExcelWorkbook workbook, List<ExcelNamedRangeSnapshot> namedRanges) {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    Set<String> workbookSheetNames = new LinkedHashSet<>(workbook.sheetNames());
    for (ExcelNamedRangeSnapshot namedRange : namedRanges) {
      WorkbookAnalysis.AnalysisLocation.NamedRange location = namedRangeLocation(namedRange);
      String refersToFormula = namedRange.refersToFormula();
      String upperFormula = refersToFormula.toUpperCase(Locale.ROOT);
      if (upperFormula.contains("#REF!")) {
        findings.add(
            new WorkbookAnalysis.AnalysisFinding(
                WorkbookAnalysis.AnalysisFindingCode.NAMED_RANGE_BROKEN_REFERENCE,
                WorkbookAnalysis.AnalysisSeverity.ERROR,
                "Named range contains a broken reference",
                "Named range formula contains #REF! and is broken.",
                location,
                List.of(refersToFormula)));
        continue;
      }

      String referencedSheet = referencedSheetName(refersToFormula);
      if (referencedSheet != null && !workbookSheetNames.contains(referencedSheet)) {
        findings.add(
            new WorkbookAnalysis.AnalysisFinding(
                WorkbookAnalysis.AnalysisFindingCode.NAMED_RANGE_BROKEN_REFERENCE,
                WorkbookAnalysis.AnalysisSeverity.ERROR,
                "Named range targets a missing sheet",
                "Named range refers to a sheet that does not exist: " + referencedSheet,
                location,
                List.of(refersToFormula)));
        continue;
      }

      if (namedRange instanceof ExcelNamedRangeSnapshot.FormulaSnapshot
          && looksLikeSheetRangeReference(refersToFormula)) {
        findings.add(
            new WorkbookAnalysis.AnalysisFinding(
                WorkbookAnalysis.AnalysisFindingCode.NAMED_RANGE_UNRESOLVED_TARGET,
                WorkbookAnalysis.AnalysisSeverity.WARNING,
                "Named range target could not be normalized",
                "Named range looks like a sheet-qualified reference but did not normalize cleanly.",
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
                WorkbookAnalysis.AnalysisFindingCode.NAMED_RANGE_SCOPE_SHADOWING,
                WorkbookAnalysis.AnalysisSeverity.INFO,
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
      case ExcelSheetSelection.All _ -> workbook.sheetNames();
      case ExcelSheetSelection.Selected selected -> List.copyOf(selected.sheetNames());
    };
  }

  WorkbookAnalysis.AnalysisSummary summary(List<WorkbookAnalysis.AnalysisFinding> findings) {
    int errorCount = 0;
    int warningCount = 0;
    int infoCount = 0;
    for (WorkbookAnalysis.AnalysisFinding finding : findings) {
      if (finding.severity() == WorkbookAnalysis.AnalysisSeverity.ERROR) {
        errorCount++;
        continue;
      }
      if (finding.severity() == WorkbookAnalysis.AnalysisSeverity.WARNING) {
        warningCount++;
        continue;
      }
      infoCount++;
    }
    return new WorkbookAnalysis.AnalysisSummary(
        findings.size(), errorCount, warningCount, infoCount);
  }

  private static WorkbookAnalysis.AnalysisLocation.NamedRange namedRangeLocation(
      ExcelNamedRangeSnapshot namedRange) {
    return new WorkbookAnalysis.AnalysisLocation.NamedRange(namedRange.name(), namedRange.scope());
  }

  String referencedSheetName(String refersToFormula) {
    int bangIndex = refersToFormula.indexOf('!');
    if (bangIndex <= 0) {
      return null;
    }
    String sheetName = refersToFormula.substring(0, bangIndex);
    if (!sheetName.startsWith("'")) {
      return sheetName;
    }
    if (!sheetName.endsWith("'")) {
      return sheetName;
    }
    if (sheetName.length() < 2) {
      return sheetName;
    }
    return sheetName.substring(1, sheetName.length() - 1).replace("''", "'");
  }

  boolean looksLikeSheetRangeReference(String formula) {
    int bangIndex = formula.indexOf('!');
    if (bangIndex <= 0 || bangIndex >= formula.length() - 1) {
      return false;
    }
    String tail = formula.substring(bangIndex + 1);
    return tail.contains("$") || tail.contains(":") || Character.isLetter(tail.charAt(0));
  }
}
