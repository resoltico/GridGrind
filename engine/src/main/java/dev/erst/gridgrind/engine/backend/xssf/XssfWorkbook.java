package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.result.SheetSummaryReport;
import dev.erst.gridgrind.core.result.WorkbookSummary;
import dev.erst.gridgrind.engine.customxml.CustomXmlAdapter;
import dev.erst.gridgrind.engine.formulas.FormulaAdapter;
import dev.erst.gridgrind.engine.formulas.FormulaContext;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import dev.erst.gridgrind.engine.names.NamedRangeAdapter;
import dev.erst.gridgrind.engine.protection.ProtectionAdapter;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * High-level workbook wrapper around Apache POI for creation, loading, saving, and sheet access.
 */
public final class XssfWorkbook implements AutoCloseable {
  private final WorkbookContext context;
  private final FormulaAdapter formulas;
  private final SheetCollectionAdapter sheets;
  private final CustomXmlAdapter customXml;
  private final ProtectionAdapter protection;
  private final NamedRangeAdapter names;
  private final WorkbookPersistence persistence;
  private boolean mutatedSinceOpen;

  XssfWorkbook(XSSFWorkbook workbook) {
    this(
        workbook,
        FormulaRuntime.poi(workbook.getCreationHelper().createFormulaEvaluator()),
        Optional.empty(),
        OoxmlPackageSecurityReport.none(),
        Optional.empty());
  }

  XssfWorkbook(XSSFWorkbook workbook, ResolvedFormulaEnvironment formulaEnvironment)
      throws IOException {
    this(
        workbook,
        FormulaRuntime.poi(workbook, formulaEnvironment),
        Optional.empty(),
        OoxmlPackageSecurityReport.none(),
        Optional.empty());
  }

  XssfWorkbook(XSSFWorkbook workbook, FormulaRuntime formulaRuntime) {
    this(
        workbook,
        formulaRuntime,
        Optional.empty(),
        OoxmlPackageSecurityReport.none(),
        Optional.empty());
  }

  XssfWorkbook(
      XSSFWorkbook workbook,
      FormulaRuntime formulaRuntime,
      Optional<Path> sourcePath,
      OoxmlPackageSecurityReport loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword) {
    this.context =
        new WorkbookContext(
            workbook, formulaRuntime, sourcePath, loadedPackageSecurity, sourceEncryptionPassword);
    this.formulas = new FormulaAdapter(workbook, formulaRuntime, this::markPackageMutated);
    this.sheets = new SheetCollectionAdapter(this);
    this.customXml = new CustomXmlAdapter(xssfWorkbook());
    this.protection = new ProtectionAdapter(xssfWorkbook());
    this.names = new NamedRangeAdapter(this);
    this.persistence = new WorkbookPersistence(this);
  }

  /** Adapts a POI evaluator into the GridGrind-owned formula runtime seam. */
  XssfWorkbook(XSSFWorkbook workbook, FormulaEvaluator formulaEvaluator) {
    this(workbook, FormulaRuntime.poi(formulaEvaluator));
  }

  /** Returns the formula-operation surface for evaluation, cache management, and diagnostics. */
  public FormulaAdapter formulas() {
    return formulas;
  }

  /** Returns the grouped sheet lifecycle, visibility, and protection surface. */
  public SheetCollectionAdapter sheets() {
    return sheets;
  }

  /** Returns the workbook custom-XML mapping surface. */
  public CustomXmlAdapter customXml() {
    return customXml;
  }

  /** Returns the workbook-level protection surface. */
  public ProtectionAdapter protection() {
    return protection;
  }

  /** Returns the workbook defined-name authoring and inspection surface. */
  public NamedRangeAdapter names() {
    return names;
  }

  /** Returns the workbook persistence and source-metadata surface. */
  public WorkbookPersistence persistence() {
    return persistence;
  }

  /** Returns the named sheet, creating it if necessary. */
  public XssfSheet getOrCreateSheet(String sheetName) {
    return sheets().getOrCreateSheet(sheetName);
  }

  /** Returns an existing sheet. */
  public XssfSheet sheet(String sheetName) {
    return sheets().sheet(sheetName);
  }

  /** Resets only the in-process evaluator cache after workbook mutations. */
  void invalidateFormulaRuntime() {
    context.formulaRuntime().clearCachedResults();
  }

  /** Returns the evaluator environment facts used by formula reads and diagnostics. */
  FormulaContext formulaRuntimeContext() {
    return context.formulaRuntime().context();
  }

  /** Returns the current workbook-wide summary. */
  WorkbookSummary workbookSummary() {

    if (sheets().sheetCount() == 0) {
      return new WorkbookSummary.Empty(
          0, List.of(), names().namedRangeCount(), formulas().recalculateOnOpenEnabled());
    }
    return new WorkbookSummary.WithSheets(
        sheets().sheetCount(),
        sheets().sheetNames().stream()
            .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
            .toList(),
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(
            SheetCollectionAdapter.activeSheetName(xssfWorkbook())),
        SheetCollectionAdapter.selectedSheetNames(xssfWorkbook()).stream()
            .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
            .toList(),
        names().namedRangeCount(),
        formulas().recalculateOnOpenEnabled());
  }

  /** Returns the summary facts for one sheet, including visibility and protection state. */
  SheetSummaryReport sheetSummary(String sheetName) {
    return sheets().summarizeSheet(sheetName);
  }

  /** Returns factual OOXML package-security state for the current in-memory workbook. */
  OoxmlPackageSecurityReport packageSecurity() {
    return persistence.packageSecurity();
  }

  @Override
  public void close() throws IOException {
    persistence.close();
  }

  /** Returns the mutable XSSF workbook delegate used by workbook-scoped controllers. */
  public XSSFWorkbook xssfWorkbook() {
    return context.workbook();
  }

  WorkbookContext context() {
    return context;
  }

  boolean wasMutatedSinceOpenInternal() {
    return mutatedSinceOpen;
  }

  void markPackageMutated() {
    mutatedSinceOpen = true;
  }

  int requiredSheetIndex(String sheetName) {
    return SheetCollectionAdapter.requiredSheetIndex(xssfWorkbook(), sheetName);
  }
}
